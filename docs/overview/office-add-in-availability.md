---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 11/07/2018
ms.openlocfilehash: 9490fca9663737e2397de159169b545e3900289f
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458042"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="e7594-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="e7594-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="e7594-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="e7594-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、各 Office アプリケーションで現在サポートされている共通 API が記載されています。</span><span class="sxs-lookup"><span data-stu-id="e7594-105">The following tables contain the available platforms, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="e7594-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="e7594-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="e7594-108">Excel</span><span class="sxs-lookup"><span data-stu-id="e7594-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="e7594-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="e7594-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="e7594-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="e7594-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="e7594-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e7594-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="e7594-113">Office Online</span></span></td>
    <td> <span data-ttu-id="e7594-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-114">- TaskPane</span></span><br><span data-ttu-id="e7594-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-115">
        - Content</span></span><br><span data-ttu-id="e7594-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="e7594-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="e7594-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e7594-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e7594-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e7594-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e7594-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e7594-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e7594-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e7594-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e7594-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e7594-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-126">
        - BindingEvents</span></span><br><span data-ttu-id="e7594-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-127">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-128">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-129">
        - File</span></span><br><span data-ttu-id="e7594-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-130">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-132">
        - Selection</span></span><br><span data-ttu-id="e7594-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-133">
        - Settings</span></span><br><span data-ttu-id="e7594-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-134">
        - TableBindings</span></span><br><span data-ttu-id="e7594-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-135">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-136">
        - TextBindings</span></span><br><span data-ttu-id="e7594-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="e7594-139">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-139">
        - TaskPane</span></span><br><span data-ttu-id="e7594-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="e7594-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-142">
        - BindingEvents</span></span><br><span data-ttu-id="e7594-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-143">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-144">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-145">
        - File</span></span><br><span data-ttu-id="e7594-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-146">
        - ImageCoercion</span></span><br><span data-ttu-id="e7594-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-147">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-149">
        - Selection</span></span><br><span data-ttu-id="e7594-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-150">
        - Settings</span></span><br><span data-ttu-id="e7594-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-151">
        - TableBindings</span></span><br><span data-ttu-id="e7594-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-152">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-153">
        - TextBindings</span></span><br><span data-ttu-id="e7594-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="e7594-156">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-156">- TaskPane</span></span><br><span data-ttu-id="e7594-157">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-157">
        - Content</span></span><br><span data-ttu-id="e7594-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e7594-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e7594-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e7594-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e7594-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e7594-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e7594-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e7594-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e7594-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e7594-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e7594-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-168">- BindingEvents</span></span><br><span data-ttu-id="e7594-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-169">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-170">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-171">
        - File</span></span><br><span data-ttu-id="e7594-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-172">
        - ImageCoercion</span></span><br><span data-ttu-id="e7594-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-173">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-175">
        - Selection</span></span><br><span data-ttu-id="e7594-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-176">
        - Settings</span></span><br><span data-ttu-id="e7594-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-177">
        - TableBindings</span></span><br><span data-ttu-id="e7594-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-178">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-179">
        - TextBindings</span></span><br><span data-ttu-id="e7594-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="e7594-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-182">- TaskPane</span></span><br><span data-ttu-id="e7594-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-183">
        - Content</span></span><br><span data-ttu-id="e7594-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e7594-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e7594-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e7594-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e7594-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e7594-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e7594-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e7594-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e7594-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e7594-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e7594-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-194">- BindingEvents</span></span><br><span data-ttu-id="e7594-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-195">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-196">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-197">
        - File</span></span><br><span data-ttu-id="e7594-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-198">
        - ImageCoercion</span></span><br><span data-ttu-id="e7594-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-199">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-201">
        - Selection</span></span><br><span data-ttu-id="e7594-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-202">
        - Settings</span></span><br><span data-ttu-id="e7594-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-203">
        - TableBindings</span></span><br><span data-ttu-id="e7594-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-204">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-205">
        - TextBindings</span></span><br><span data-ttu-id="e7594-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="e7594-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="e7594-208">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-208">- TaskPane</span></span><br><span data-ttu-id="e7594-209">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-209">
        - Content</span></span></td>
    <td><span data-ttu-id="e7594-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e7594-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e7594-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e7594-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e7594-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e7594-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e7594-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e7594-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e7594-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e7594-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-219">- BindingEvents</span></span><br><span data-ttu-id="e7594-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-220">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-221">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-222">
        - File</span></span><br><span data-ttu-id="e7594-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-223">
        - ImageCoercion</span></span><br><span data-ttu-id="e7594-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-224">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-226">
        - Selection</span></span><br><span data-ttu-id="e7594-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-227">
        - Settings</span></span><br><span data-ttu-id="e7594-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-228">
        - TableBindings</span></span><br><span data-ttu-id="e7594-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-229">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-230">
        - TextBindings</span></span><br><span data-ttu-id="e7594-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="e7594-233">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-233">- TaskPane</span></span><br><span data-ttu-id="e7594-234">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-234">
        - Content</span></span><br><span data-ttu-id="e7594-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e7594-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e7594-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e7594-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e7594-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e7594-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e7594-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e7594-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e7594-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e7594-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e7594-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-245">- BindingEvents</span></span><br><span data-ttu-id="e7594-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-246">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-247">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-248">
        - File</span></span><br><span data-ttu-id="e7594-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-249">
        - ImageCoercion</span></span><br><span data-ttu-id="e7594-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-250">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-252">
        - PdfFile</span></span><br><span data-ttu-id="e7594-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-253">
        - Selection</span></span><br><span data-ttu-id="e7594-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-254">
        - Settings</span></span><br><span data-ttu-id="e7594-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-255">
        - TableBindings</span></span><br><span data-ttu-id="e7594-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-256">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-257">
        - TextBindings</span></span><br><span data-ttu-id="e7594-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="e7594-260">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-260">- TaskPane</span></span><br><span data-ttu-id="e7594-261">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-261">
        - Content</span></span><br><span data-ttu-id="e7594-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="e7594-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="e7594-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="e7594-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="e7594-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="e7594-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="e7594-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="e7594-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="e7594-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="e7594-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="e7594-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="e7594-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-272">- BindingEvents</span></span><br><span data-ttu-id="e7594-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-273">
        - CompressedFile</span></span><br><span data-ttu-id="e7594-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-274">
        - DocumentEvents</span></span><br><span data-ttu-id="e7594-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="e7594-275">
        - File</span></span><br><span data-ttu-id="e7594-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-276">
        - ImageCoercion</span></span><br><span data-ttu-id="e7594-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-277">
        - MatrixBindings</span></span><br><span data-ttu-id="e7594-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="e7594-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-279">
        - PdfFile</span></span><br><span data-ttu-id="e7594-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-280">
        - Selection</span></span><br><span data-ttu-id="e7594-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-281">
        - Settings</span></span><br><span data-ttu-id="e7594-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-282">
        - TableBindings</span></span><br><span data-ttu-id="e7594-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-283">
        - TableCoercion</span></span><br><span data-ttu-id="e7594-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-284">
        - TextBindings</span></span><br><span data-ttu-id="e7594-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="e7594-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="e7594-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e7594-287">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-287">Platform</span></span></th>
    <th><span data-ttu-id="e7594-288">拡張点</span><span class="sxs-lookup"><span data-stu-id="e7594-288">Extension points</span></span></th>
    <th><span data-ttu-id="e7594-289">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="e7594-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e7594-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="e7594-291">Office Online</span></span></td>
    <td> <span data-ttu-id="e7594-292">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-292">- Mail Read</span></span><br><span data-ttu-id="e7594-293">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e7594-293">
      - Mail Compose</span></span><br><span data-ttu-id="e7594-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e7594-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e7594-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e7594-302">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-304">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-304">- Mail Read</span></span><br><span data-ttu-id="e7594-305">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e7594-305">
      - Mail Compose</span></span><br><span data-ttu-id="e7594-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="e7594-311">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-313">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-313">- Mail Read</span></span><br><span data-ttu-id="e7594-314">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e7594-314">
      - Mail Compose</span></span><br><span data-ttu-id="e7594-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e7594-316">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="e7594-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e7594-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e7594-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e7594-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e7594-324">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-326">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-326">- Mail Read</span></span><br><span data-ttu-id="e7594-327">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e7594-327">
      - Mail Compose</span></span><br><span data-ttu-id="e7594-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="e7594-329">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="e7594-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="e7594-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e7594-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e7594-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e7594-337">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="e7594-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="e7594-339">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-339">- Mail Read</span></span><br><span data-ttu-id="e7594-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e7594-346">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e7594-348">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-348">- Mail Read</span></span><br><span data-ttu-id="e7594-349">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e7594-349">
      - Mail Compose</span></span><br><span data-ttu-id="e7594-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e7594-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="e7594-357">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e7594-359">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-359">- Mail Read</span></span><br><span data-ttu-id="e7594-360">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="e7594-360">
      - Mail Compose</span></span><br><span data-ttu-id="e7594-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="e7594-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="e7594-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="e7594-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="e7594-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="e7594-369">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="e7594-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="e7594-371">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="e7594-371">- Mail Read</span></span><br><span data-ttu-id="e7594-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="e7594-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="e7594-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="e7594-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="e7594-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="e7594-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="e7594-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="e7594-378">利用不可</span><span class="sxs-lookup"><span data-stu-id="e7594-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="e7594-379">Word</span><span class="sxs-lookup"><span data-stu-id="e7594-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e7594-380">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-380">Platform</span></span></th>
    <th><span data-ttu-id="e7594-381">拡張点</span><span class="sxs-lookup"><span data-stu-id="e7594-381">Extension points</span></span></th>
    <th><span data-ttu-id="e7594-382">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="e7594-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e7594-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="e7594-384">Office Online</span></span></td>
    <td> <span data-ttu-id="e7594-385">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-385">- TaskPane</span></span><br><span data-ttu-id="e7594-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e7594-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e7594-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e7594-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-391">- BindingEvents</span></span><br><span data-ttu-id="e7594-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-393">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-394">
         - File</span></span><br><span data-ttu-id="e7594-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-396">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-397">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-400">
         - PdfFile</span></span><br><span data-ttu-id="e7594-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-401">
         - Selection</span></span><br><span data-ttu-id="e7594-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-402">
         - Settings</span></span><br><span data-ttu-id="e7594-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-403">
         - TableBindings</span></span><br><span data-ttu-id="e7594-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-404">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-405">
         - TextBindings</span></span><br><span data-ttu-id="e7594-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-406">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-409">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e7594-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-411">- BindingEvents</span></span><br><span data-ttu-id="e7594-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-412">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-414">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-415">
         - File</span></span><br><span data-ttu-id="e7594-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-417">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-418">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-421">
         - PdfFile</span></span><br><span data-ttu-id="e7594-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-422">
         - Selection</span></span><br><span data-ttu-id="e7594-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-423">
         - Settings</span></span><br><span data-ttu-id="e7594-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-424">
         - TableBindings</span></span><br><span data-ttu-id="e7594-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-425">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-426">
         - TextBindings</span></span><br><span data-ttu-id="e7594-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-427">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-430">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-430">- TaskPane</span></span><br><span data-ttu-id="e7594-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e7594-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e7594-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e7594-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-436">- BindingEvents</span></span><br><span data-ttu-id="e7594-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-437">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-439">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-440">
         - File</span></span><br><span data-ttu-id="e7594-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-442">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-443">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-446">
         - PdfFile</span></span><br><span data-ttu-id="e7594-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-447">
         - Selection</span></span><br><span data-ttu-id="e7594-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-448">
         - Settings</span></span><br><span data-ttu-id="e7594-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-449">
         - TableBindings</span></span><br><span data-ttu-id="e7594-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-450">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-451">
         - TextBindings</span></span><br><span data-ttu-id="e7594-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-452">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-455">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-455">- TaskPane</span></span><br><span data-ttu-id="e7594-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e7594-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e7594-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e7594-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-461">- BindingEvents</span></span><br><span data-ttu-id="e7594-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-462">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-464">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-465">
         - File</span></span><br><span data-ttu-id="e7594-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-467">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-468">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-471">
         - PdfFile</span></span><br><span data-ttu-id="e7594-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-472">
         - Selection</span></span><br><span data-ttu-id="e7594-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-473">
         - Settings</span></span><br><span data-ttu-id="e7594-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-474">
         - TableBindings</span></span><br><span data-ttu-id="e7594-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-475">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-476">
         - TextBindings</span></span><br><span data-ttu-id="e7594-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-477">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-479">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="e7594-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e7594-480">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e7594-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e7594-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e7594-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e7594-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e7594-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e7594-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-485">- BindingEvents</span></span><br><span data-ttu-id="e7594-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-486">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-488">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-489">
         - File</span></span><br><span data-ttu-id="e7594-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-491">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-492">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-495">
         - PdfFile</span></span><br><span data-ttu-id="e7594-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-496">
         - Selection</span></span><br><span data-ttu-id="e7594-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-497">
         - Settings</span></span><br><span data-ttu-id="e7594-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-498">
         - TableBindings</span></span><br><span data-ttu-id="e7594-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-499">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-500">
         - TextBindings</span></span><br><span data-ttu-id="e7594-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-501">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e7594-504">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-504">- TaskPane</span></span><br><span data-ttu-id="e7594-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e7594-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e7594-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e7594-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e7594-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e7594-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-510">- BindingEvents</span></span><br><span data-ttu-id="e7594-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-511">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-513">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-514">
         - File</span></span><br><span data-ttu-id="e7594-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-516">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-517">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-520">
         - PdfFile</span></span><br><span data-ttu-id="e7594-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-521">
         - Selection</span></span><br><span data-ttu-id="e7594-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-522">
         - Settings</span></span><br><span data-ttu-id="e7594-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-523">
         - TableBindings</span></span><br><span data-ttu-id="e7594-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-524">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-525">
         - TextBindings</span></span><br><span data-ttu-id="e7594-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-526">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e7594-529">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-529">- TaskPane</span></span><br><span data-ttu-id="e7594-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="e7594-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="e7594-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="e7594-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="e7594-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="e7594-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e7594-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e7594-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-535">- BindingEvents</span></span><br><span data-ttu-id="e7594-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-536">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="e7594-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="e7594-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-538">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-539">
         - File</span></span><br><span data-ttu-id="e7594-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-541">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-542">
         - MatrixBindings</span></span><br><span data-ttu-id="e7594-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="e7594-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="e7594-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-545">
         - PdfFile</span></span><br><span data-ttu-id="e7594-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-546">
         - Selection</span></span><br><span data-ttu-id="e7594-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-547">
         - Settings</span></span><br><span data-ttu-id="e7594-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-548">
         - TableBindings</span></span><br><span data-ttu-id="e7594-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-549">
         - TableCoercion</span></span><br><span data-ttu-id="e7594-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="e7594-550">
         - TextBindings</span></span><br><span data-ttu-id="e7594-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-551">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="e7594-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="e7594-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="e7594-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e7594-554">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-554">Platform</span></span></th>
    <th><span data-ttu-id="e7594-555">拡張点</span><span class="sxs-lookup"><span data-stu-id="e7594-555">Extension points</span></span></th>
    <th><span data-ttu-id="e7594-556">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="e7594-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e7594-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="e7594-558">Office Online</span></span></td>
    <td> <span data-ttu-id="e7594-559">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-559">- Content</span></span><br><span data-ttu-id="e7594-560">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-560">
         - TaskPane</span></span><br><span data-ttu-id="e7594-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-563">- ActiveView</span></span><br><span data-ttu-id="e7594-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-564">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-565">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-566">
         - File</span></span><br><span data-ttu-id="e7594-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-567">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-568">
         - PdfFile</span></span><br><span data-ttu-id="e7594-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-569">
         - Selection</span></span><br><span data-ttu-id="e7594-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-570">
         - Settings</span></span><br><span data-ttu-id="e7594-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-573">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-573">- Content</span></span><br><span data-ttu-id="e7594-574">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="e7594-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="e7594-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="e7594-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-576">- ActiveView</span></span><br><span data-ttu-id="e7594-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-577">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-578">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-579">
         - File</span></span><br><span data-ttu-id="e7594-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-580">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-581">
         - PdfFile</span></span><br><span data-ttu-id="e7594-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-582">
         - Selection</span></span><br><span data-ttu-id="e7594-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-583">
         - Settings</span></span><br><span data-ttu-id="e7594-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-586">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-586">- Content</span></span><br><span data-ttu-id="e7594-587">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-587">
         - TaskPane</span></span><br><span data-ttu-id="e7594-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-590">- ActiveView</span></span><br><span data-ttu-id="e7594-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-591">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-592">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-593">
         - File</span></span><br><span data-ttu-id="e7594-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-594">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-595">
         - PdfFile</span></span><br><span data-ttu-id="e7594-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-596">
         - Selection</span></span><br><span data-ttu-id="e7594-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-597">
         - Settings</span></span><br><span data-ttu-id="e7594-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-600">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-600">- Content</span></span><br><span data-ttu-id="e7594-601">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-601">
         - TaskPane</span></span><br><span data-ttu-id="e7594-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-604">- ActiveView</span></span><br><span data-ttu-id="e7594-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-605">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-606">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-607">
         - File</span></span><br><span data-ttu-id="e7594-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-608">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-609">
         - PdfFile</span></span><br><span data-ttu-id="e7594-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-610">
         - Selection</span></span><br><span data-ttu-id="e7594-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-611">
         - Settings</span></span><br><span data-ttu-id="e7594-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-613">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="e7594-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="e7594-614">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-614">- Content</span></span><br><span data-ttu-id="e7594-615">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="e7594-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="e7594-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-617">- ActiveView</span></span><br><span data-ttu-id="e7594-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-618">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-619">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-620">
         - File</span></span><br><span data-ttu-id="e7594-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-621">
         - PdfFile</span></span><br><span data-ttu-id="e7594-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-622">
         - Selection</span></span><br><span data-ttu-id="e7594-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-623">
         - Settings</span></span><br><span data-ttu-id="e7594-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-624">
         - TextCoercion</span></span><br><span data-ttu-id="e7594-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="e7594-627">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-627">- Content</span></span><br><span data-ttu-id="e7594-628">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-628">
         - TaskPane</span></span><br><span data-ttu-id="e7594-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-631">- ActiveView</span></span><br><span data-ttu-id="e7594-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-632">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-633">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-634">
         - File</span></span><br><span data-ttu-id="e7594-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-635">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-636">
         - PdfFile</span></span><br><span data-ttu-id="e7594-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-637">
         - Selection</span></span><br><span data-ttu-id="e7594-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-638">
         - Settings</span></span><br><span data-ttu-id="e7594-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="e7594-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="e7594-641">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-641">- Content</span></span><br><span data-ttu-id="e7594-642">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-642">
         - TaskPane</span></span><br><span data-ttu-id="e7594-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="e7594-645">- ActiveView</span></span><br><span data-ttu-id="e7594-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="e7594-646">
         - CompressedFile</span></span><br><span data-ttu-id="e7594-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-647">
         - DocumentEvents</span></span><br><span data-ttu-id="e7594-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="e7594-648">
         - File</span></span><br><span data-ttu-id="e7594-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-649">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="e7594-650">
         - PdfFile</span></span><br><span data-ttu-id="e7594-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-651">
         - Selection</span></span><br><span data-ttu-id="e7594-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-652">
         - Settings</span></span><br><span data-ttu-id="e7594-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="e7594-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="e7594-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e7594-655">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-655">Platform</span></span></th>
    <th><span data-ttu-id="e7594-656">拡張点</span><span class="sxs-lookup"><span data-stu-id="e7594-656">Extension points</span></span></th>
    <th><span data-ttu-id="e7594-657">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="e7594-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e7594-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="e7594-659">Office Online</span></span></td>
    <td> <span data-ttu-id="e7594-660">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="e7594-660">- Content</span></span><br><span data-ttu-id="e7594-661">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-661">
         - TaskPane</span></span><br><span data-ttu-id="e7594-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="e7594-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="e7594-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="e7594-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="e7594-665">- DocumentEvents</span></span><br><span data-ttu-id="e7594-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="e7594-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-667">
         - ImageCoercion</span></span><br><span data-ttu-id="e7594-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="e7594-668">
         - Settings</span></span><br><span data-ttu-id="e7594-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="e7594-670">Project</span><span class="sxs-lookup"><span data-stu-id="e7594-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="e7594-671">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="e7594-671">Platform</span></span></th>
    <th><span data-ttu-id="e7594-672">拡張点</span><span class="sxs-lookup"><span data-stu-id="e7594-672">Extension points</span></span></th>
    <th><span data-ttu-id="e7594-673">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="e7594-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="e7594-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-676">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e7594-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-678">- Selection</span></span><br><span data-ttu-id="e7594-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-681">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e7594-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-683">- Selection</span></span><br><span data-ttu-id="e7594-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="e7594-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="e7594-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="e7594-686">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="e7594-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="e7594-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="e7594-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="e7594-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="e7594-688">- Selection</span></span><br><span data-ttu-id="e7594-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="e7594-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="e7594-690">関連項目</span><span class="sxs-lookup"><span data-stu-id="e7594-690">See also</span></span>

- [<span data-ttu-id="e7594-691">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="e7594-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="e7594-692">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="e7594-693">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="e7594-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="e7594-694">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="e7594-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
