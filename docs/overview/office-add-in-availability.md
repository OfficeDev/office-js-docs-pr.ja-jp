---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 11/07/2018
ms.openlocfilehash: c601eac5ed3fcad76b63fff5ae6eeadb7662c8b7
ms.sourcegitcommit: 0adc31ceaba92cb15dc6430c00fe7a96c107c9de
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/09/2018
ms.locfileid: "27210106"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="bfc6f-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="bfc6f-104">Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。</span><span class="sxs-lookup"><span data-stu-id="bfc6f-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="bfc6f-105">次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API の要件セットが含まれています。</span><span class="sxs-lookup"><span data-stu-id="bfc6f-105">The following tables contain the available platform, extension points, API requirement sets, and common API requirement sets that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="bfc6f-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、および共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="bfc6f-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and common API requirement sets.</span></span>

## <a name="excel"></a><span data-ttu-id="bfc6f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="bfc6f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="bfc6f-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="bfc6f-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc6f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="bfc6f-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="bfc6f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-112"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="bfc6f-113">Office Online</span></span></td>
    <td> <span data-ttu-id="bfc6f-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-114">- TaskPane</span></span><br><span data-ttu-id="bfc6f-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-115">
        - Content</span></span><br><span data-ttu-id="bfc6f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="bfc6f-116">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="bfc6f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bfc6f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bfc6f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bfc6f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bfc6f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bfc6f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-126">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-126">
        - BindingEvents</span></span><br><span data-ttu-id="bfc6f-127">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-127">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-128">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-128">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-129">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-129">
        - File</span></span><br><span data-ttu-id="bfc6f-130">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-130">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-131">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-131">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-132">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-132">
        - Selection</span></span><br><span data-ttu-id="bfc6f-133">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-133">
        - Settings</span></span><br><span data-ttu-id="bfc6f-134">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-134">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-135">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-135">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-136">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-136">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-137">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-137">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-138">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-138">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="bfc6f-139">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-139">
        - TaskPane</span></span><br><span data-ttu-id="bfc6f-140">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-140">
        - Content</span></span></td>
    <td>  <span data-ttu-id="bfc6f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-141">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-142">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-142">
        - BindingEvents</span></span><br><span data-ttu-id="bfc6f-143">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-143">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-144">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-144">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-145">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-145">
        - File</span></span><br><span data-ttu-id="bfc6f-146">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-146">
        - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-147">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-147">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-148">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-148">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-149">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-149">
        - Selection</span></span><br><span data-ttu-id="bfc6f-150">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-150">
        - Settings</span></span><br><span data-ttu-id="bfc6f-151">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-151">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-152">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-152">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-153">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-153">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-154">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-154">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-155">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-155">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="bfc6f-156">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-156">- TaskPane</span></span><br><span data-ttu-id="bfc6f-157">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-157">
        - Content</span></span><br><span data-ttu-id="bfc6f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-158">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bfc6f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-159">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-160">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-161">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-162">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bfc6f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-163">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bfc6f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-164">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bfc6f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-165">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bfc6f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-166">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bfc6f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-168">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-168">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-169">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-169">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-170">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-170">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-171">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-171">
        - File</span></span><br><span data-ttu-id="bfc6f-172">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-172">
        - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-173">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-173">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-174">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-174">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-175">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-175">
        - Selection</span></span><br><span data-ttu-id="bfc6f-176">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-176">
        - Settings</span></span><br><span data-ttu-id="bfc6f-177">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-177">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-178">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-178">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-179">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-179">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-180">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-180">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-181">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-181">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="bfc6f-182">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-182">- TaskPane</span></span><br><span data-ttu-id="bfc6f-183">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-183">
        - Content</span></span><br><span data-ttu-id="bfc6f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-184">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bfc6f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-185">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-186">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-187">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-188">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bfc6f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-189">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bfc6f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-190">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bfc6f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-191">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bfc6f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-192">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bfc6f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-193">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-194">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-194">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-195">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-195">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-196">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-196">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-197">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-197">
        - File</span></span><br><span data-ttu-id="bfc6f-198">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-198">
        - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-199">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-199">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-200">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-200">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-201">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-201">
        - Selection</span></span><br><span data-ttu-id="bfc6f-202">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-202">
        - Settings</span></span><br><span data-ttu-id="bfc6f-203">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-203">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-204">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-204">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-205">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-205">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-206">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-206">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-207">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="bfc6f-207">Office for iPad</span></span></td>
    <td><span data-ttu-id="bfc6f-208">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-208">- TaskPane</span></span><br><span data-ttu-id="bfc6f-209">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-209">
        - Content</span></span></td>
    <td><span data-ttu-id="bfc6f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-210">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-211">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-212">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-213">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bfc6f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-214">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bfc6f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-215">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bfc6f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-216">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bfc6f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-217">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bfc6f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-218">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-219">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-219">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-220">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-221">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-222">
        - File</span></span><br><span data-ttu-id="bfc6f-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-223">
        - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-224">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-226">
        - Selection</span></span><br><span data-ttu-id="bfc6f-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-227">
        - Settings</span></span><br><span data-ttu-id="bfc6f-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-228">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-229">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-230">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-232">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-232">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="bfc6f-233">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-233">- TaskPane</span></span><br><span data-ttu-id="bfc6f-234">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-234">
        - Content</span></span><br><span data-ttu-id="bfc6f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-235">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bfc6f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-236">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-237">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-238">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-239">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bfc6f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-240">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bfc6f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-241">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bfc6f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-242">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bfc6f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-243">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bfc6f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-244">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-245">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-245">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-246">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-246">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-247">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-247">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-248">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-248">
        - File</span></span><br><span data-ttu-id="bfc6f-249">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-249">
        - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-250">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-250">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-251">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-251">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-252">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-252">
        - PdfFile</span></span><br><span data-ttu-id="bfc6f-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-253">
        - Selection</span></span><br><span data-ttu-id="bfc6f-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-254">
        - Settings</span></span><br><span data-ttu-id="bfc6f-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-255">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-256">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-257">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-259">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-259">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="bfc6f-260">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-260">- TaskPane</span></span><br><span data-ttu-id="bfc6f-261">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-261">
        - Content</span></span><br><span data-ttu-id="bfc6f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="bfc6f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-263">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-264">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-265">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-266">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="bfc6f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-267">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="bfc6f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-268">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="bfc6f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-269">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="bfc6f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-270">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="bfc6f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-271">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="bfc6f-272">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-272">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-273">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-273">
        - CompressedFile</span></span><br><span data-ttu-id="bfc6f-274">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-274">
        - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-275">
        - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-275">
        - File</span></span><br><span data-ttu-id="bfc6f-276">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-276">
        - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-277">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-277">
        - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-278">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-278">
        - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-279">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-279">
        - PdfFile</span></span><br><span data-ttu-id="bfc6f-280">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-280">
        - Selection</span></span><br><span data-ttu-id="bfc6f-281">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-281">
        - Settings</span></span><br><span data-ttu-id="bfc6f-282">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-282">
        - TableBindings</span></span><br><span data-ttu-id="bfc6f-283">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-283">
        - TableCoercion</span></span><br><span data-ttu-id="bfc6f-284">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-284">
        - TextBindings</span></span><br><span data-ttu-id="bfc6f-285">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-285">
        - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="outlook"></a><span data-ttu-id="bfc6f-286">Outlook</span><span class="sxs-lookup"><span data-stu-id="bfc6f-286">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bfc6f-287">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-287">Platform</span></span></th>
    <th><span data-ttu-id="bfc6f-288">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc6f-288">Extension points</span></span></th>
    <th><span data-ttu-id="bfc6f-289">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-289">API requirement sets</span></span></th>
    <th><span data-ttu-id="bfc6f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-290"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-291">Office Online</span><span class="sxs-lookup"><span data-stu-id="bfc6f-291">Office Online</span></span></td>
    <td> <span data-ttu-id="bfc6f-292">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-292">- Mail Read</span></span><br><span data-ttu-id="bfc6f-293">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="bfc6f-293">
      - Mail Compose</span></span><br><span data-ttu-id="bfc6f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-294">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-295">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-296">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-297">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-298">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-299">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bfc6f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-300">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bfc6f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-301">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bfc6f-302">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-302">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-303">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-303">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-304">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-304">- Mail Read</span></span><br><span data-ttu-id="bfc6f-305">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="bfc6f-305">
      - Mail Compose</span></span><br><span data-ttu-id="bfc6f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-306">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-308">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-309">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-310">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="bfc6f-311">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-311">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-312">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-312">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-313">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-313">- Mail Read</span></span><br><span data-ttu-id="bfc6f-314">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="bfc6f-314">
      - Mail Compose</span></span><br><span data-ttu-id="bfc6f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-315">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bfc6f-316">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="bfc6f-316">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bfc6f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-317">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-318">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-319">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-320">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-321">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bfc6f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-322">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bfc6f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-323">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bfc6f-324">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-324">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-325">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-325">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-326">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-326">- Mail Read</span></span><br><span data-ttu-id="bfc6f-327">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="bfc6f-327">
      - Mail Compose</span></span><br><span data-ttu-id="bfc6f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-328">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="bfc6f-329">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="bfc6f-329">
      - Modules</span></span></td>
    <td> <span data-ttu-id="bfc6f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-330">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-331">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-333">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bfc6f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bfc6f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bfc6f-337">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-337">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-338">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="bfc6f-338">Office for iOS</span></span></td>
    <td> <span data-ttu-id="bfc6f-339">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-339">- Mail Read</span></span><br><span data-ttu-id="bfc6f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-340">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-341">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-342">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-343">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-345">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bfc6f-346">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-346">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-347">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-347">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bfc6f-348">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-348">- Mail Read</span></span><br><span data-ttu-id="bfc6f-349">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="bfc6f-349">
      - Mail Compose</span></span><br><span data-ttu-id="bfc6f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-351">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-353">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-354">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-355">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bfc6f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-356">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="bfc6f-357">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-357">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-358">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-358">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="bfc6f-359">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-359">- Mail Read</span></span><br><span data-ttu-id="bfc6f-360">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="bfc6f-360">
      - Mail Compose</span></span><br><span data-ttu-id="bfc6f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-362">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-366">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="bfc6f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-367">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="bfc6f-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-368">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="bfc6f-369">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-369">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-370">Office for Android</span><span class="sxs-lookup"><span data-stu-id="bfc6f-370">Office for Android</span></span></td>
    <td> <span data-ttu-id="bfc6f-371">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="bfc6f-371">- Mail Read</span></span><br><span data-ttu-id="bfc6f-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-372">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-373">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="bfc6f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="bfc6f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="bfc6f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-376">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="bfc6f-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-377">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="bfc6f-378">利用不可</span><span class="sxs-lookup"><span data-stu-id="bfc6f-378">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="bfc6f-379">Word</span><span class="sxs-lookup"><span data-stu-id="bfc6f-379">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bfc6f-380">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-380">Platform</span></span></th>
    <th><span data-ttu-id="bfc6f-381">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc6f-381">Extension points</span></span></th>
    <th><span data-ttu-id="bfc6f-382">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-382">API requirement sets</span></span></th>
    <th><span data-ttu-id="bfc6f-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-383"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-384">Office Online</span><span class="sxs-lookup"><span data-stu-id="bfc6f-384">Office Online</span></span></td>
    <td> <span data-ttu-id="bfc6f-385">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-385">- TaskPane</span></span><br><span data-ttu-id="bfc6f-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-386">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-387">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-388">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-389">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-390">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-391">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-391">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-392">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-392">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-393">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-393">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-394">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-394">
         - File</span></span><br><span data-ttu-id="bfc6f-395">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-395">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-396">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-396">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-397">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-397">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-398">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-398">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-399">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-399">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-400">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-400">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-401">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-401">
         - Selection</span></span><br><span data-ttu-id="bfc6f-402">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-402">
         - Settings</span></span><br><span data-ttu-id="bfc6f-403">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-403">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-404">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-404">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-405">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-405">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-406">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-406">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-407">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-407">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-408">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-408">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-409">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-409">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bfc6f-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-410">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-411">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-411">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-412">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-412">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-413">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-413">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-414">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-414">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-415">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-415">
         - File</span></span><br><span data-ttu-id="bfc6f-416">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-416">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-417">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-417">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-418">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-418">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-419">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-419">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-420">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-420">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-421">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-421">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-422">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-422">
         - Selection</span></span><br><span data-ttu-id="bfc6f-423">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-423">
         - Settings</span></span><br><span data-ttu-id="bfc6f-424">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-424">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-425">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-425">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-426">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-426">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-427">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-427">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-428">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-428">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-429">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-429">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-430">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-430">- TaskPane</span></span><br><span data-ttu-id="bfc6f-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-431">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-432">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-433">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-434">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-435">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-436">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-436">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-437">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-437">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-438">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-438">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-439">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-439">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-440">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-440">
         - File</span></span><br><span data-ttu-id="bfc6f-441">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-441">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-442">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-442">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-443">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-443">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-444">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-444">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-445">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-445">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-446">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-446">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-447">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-447">
         - Selection</span></span><br><span data-ttu-id="bfc6f-448">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-448">
         - Settings</span></span><br><span data-ttu-id="bfc6f-449">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-449">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-450">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-450">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-451">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-451">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-452">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-452">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-453">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-453">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-454">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-454">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-455">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-455">- TaskPane</span></span><br><span data-ttu-id="bfc6f-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-456">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-457">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-458">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-459">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-460">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-461">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-461">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-462">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-462">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-463">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-463">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-464">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-464">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-465">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-465">
         - File</span></span><br><span data-ttu-id="bfc6f-466">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-466">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-467">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-467">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-468">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-468">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-469">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-469">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-470">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-470">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-471">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-471">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-472">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-472">
         - Selection</span></span><br><span data-ttu-id="bfc6f-473">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-473">
         - Settings</span></span><br><span data-ttu-id="bfc6f-474">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-474">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-475">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-475">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-476">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-476">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-477">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-477">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-478">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-478">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-479">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="bfc6f-479">Office for iPad</span></span></td>
    <td> <span data-ttu-id="bfc6f-480">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-480">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bfc6f-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-481">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-482">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-483">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bfc6f-484">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bfc6f-485">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-485">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-486">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-486">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-487">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-487">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-488">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-488">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-489">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-489">
         - File</span></span><br><span data-ttu-id="bfc6f-490">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-490">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-491">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-491">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-492">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-492">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-493">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-493">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-494">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-494">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-495">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-495">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-496">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-496">
         - Selection</span></span><br><span data-ttu-id="bfc6f-497">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-497">
         - Settings</span></span><br><span data-ttu-id="bfc6f-498">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-498">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-499">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-499">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-500">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-500">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-501">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-501">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-502">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-502">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-503">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-503">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bfc6f-504">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-504">- TaskPane</span></span><br><span data-ttu-id="bfc6f-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-505">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-506">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-507">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-508">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bfc6f-509">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bfc6f-510">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-510">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-511">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-511">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-513">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-514">
         - File</span></span><br><span data-ttu-id="bfc6f-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-516">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-516">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-517">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-517">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-518">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-518">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-519">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-519">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-520">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-520">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-521">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-521">
         - Selection</span></span><br><span data-ttu-id="bfc6f-522">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-522">
         - Settings</span></span><br><span data-ttu-id="bfc6f-523">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-523">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-524">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-524">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-525">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-525">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-526">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-526">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-527">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-527">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-528">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-528">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="bfc6f-529">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-529">- TaskPane</span></span><br><span data-ttu-id="bfc6f-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-530">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-531">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-532">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="bfc6f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-533">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="bfc6f-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bfc6f-534">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bfc6f-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-535">- BindingEvents</span></span><br><span data-ttu-id="bfc6f-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-536">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="bfc6f-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="bfc6f-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-538">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-539">
         - File</span></span><br><span data-ttu-id="bfc6f-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-541">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-541">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-542">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-542">
         - MatrixBindings</span></span><br><span data-ttu-id="bfc6f-543">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-543">
         - MatrixCoercion</span></span><br><span data-ttu-id="bfc6f-544">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-544">
         - OoxmlCoercion</span></span><br><span data-ttu-id="bfc6f-545">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-545">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-546">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-546">
         - Selection</span></span><br><span data-ttu-id="bfc6f-547">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-547">
         - Settings</span></span><br><span data-ttu-id="bfc6f-548">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-548">
         - TableBindings</span></span><br><span data-ttu-id="bfc6f-549">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-549">
         - TableCoercion</span></span><br><span data-ttu-id="bfc6f-550">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-550">
         - TextBindings</span></span><br><span data-ttu-id="bfc6f-551">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-551">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-552">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-552">
         - TextFile</span></span> </td>
  </tr>
</table>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="bfc6f-553">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="bfc6f-553">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bfc6f-554">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-554">Platform</span></span></th>
    <th><span data-ttu-id="bfc6f-555">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc6f-555">Extension points</span></span></th>
    <th><span data-ttu-id="bfc6f-556">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-556">API requirement sets</span></span></th>
    <th><span data-ttu-id="bfc6f-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-557"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-558">Office Online</span><span class="sxs-lookup"><span data-stu-id="bfc6f-558">Office Online</span></span></td>
    <td> <span data-ttu-id="bfc6f-559">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-559">- Content</span></span><br><span data-ttu-id="bfc6f-560">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-560">
         - TaskPane</span></span><br><span data-ttu-id="bfc6f-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-562">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-563">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-563">- ActiveView</span></span><br><span data-ttu-id="bfc6f-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-564">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-565">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-565">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-566">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-566">
         - File</span></span><br><span data-ttu-id="bfc6f-567">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-567">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-568">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-569">
         - Selection</span></span><br><span data-ttu-id="bfc6f-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-570">
         - Settings</span></span><br><span data-ttu-id="bfc6f-571">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-571">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-572">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-572">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-573">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-573">- Content</span></span><br><span data-ttu-id="bfc6f-574">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-574">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="bfc6f-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="bfc6f-575">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="bfc6f-576">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-576">- ActiveView</span></span><br><span data-ttu-id="bfc6f-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-577">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-578">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-578">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-579">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-579">
         - File</span></span><br><span data-ttu-id="bfc6f-580">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-580">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-581">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-581">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-582">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-582">
         - Selection</span></span><br><span data-ttu-id="bfc6f-583">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-583">
         - Settings</span></span><br><span data-ttu-id="bfc6f-584">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-584">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-585">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-585">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-586">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-586">- Content</span></span><br><span data-ttu-id="bfc6f-587">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-587">
         - TaskPane</span></span><br><span data-ttu-id="bfc6f-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-588">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-589">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-590">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-590">- ActiveView</span></span><br><span data-ttu-id="bfc6f-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-591">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-592">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-593">
         - File</span></span><br><span data-ttu-id="bfc6f-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-594">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-595">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-596">
         - Selection</span></span><br><span data-ttu-id="bfc6f-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-597">
         - Settings</span></span><br><span data-ttu-id="bfc6f-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-598">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-599">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-599">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-600">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-600">- Content</span></span><br><span data-ttu-id="bfc6f-601">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-601">
         - TaskPane</span></span><br><span data-ttu-id="bfc6f-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-602">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-603">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-604">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-604">- ActiveView</span></span><br><span data-ttu-id="bfc6f-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-605">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-606">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-606">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-607">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-607">
         - File</span></span><br><span data-ttu-id="bfc6f-608">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-608">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-609">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-609">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-610">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-610">
         - Selection</span></span><br><span data-ttu-id="bfc6f-611">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-611">
         - Settings</span></span><br><span data-ttu-id="bfc6f-612">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-612">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-613">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="bfc6f-613">Office for iPad</span></span></td>
    <td> <span data-ttu-id="bfc6f-614">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-614">- Content</span></span><br><span data-ttu-id="bfc6f-615">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-615">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="bfc6f-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-616">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="bfc6f-617">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-617">- ActiveView</span></span><br><span data-ttu-id="bfc6f-618">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-618">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-619">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-619">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-620">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-620">
         - File</span></span><br><span data-ttu-id="bfc6f-621">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-621">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-622">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-622">
         - Selection</span></span><br><span data-ttu-id="bfc6f-623">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-623">
         - Settings</span></span><br><span data-ttu-id="bfc6f-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-624">
         - TextCoercion</span></span><br><span data-ttu-id="bfc6f-625">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-625">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-626">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-626">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="bfc6f-627">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-627">- Content</span></span><br><span data-ttu-id="bfc6f-628">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-628">
         - TaskPane</span></span><br><span data-ttu-id="bfc6f-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-629">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-630">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-631">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-631">- ActiveView</span></span><br><span data-ttu-id="bfc6f-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-632">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-633">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-634">
         - File</span></span><br><span data-ttu-id="bfc6f-635">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-635">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-636">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-636">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-637">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-637">
         - Selection</span></span><br><span data-ttu-id="bfc6f-638">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-638">
         - Settings</span></span><br><span data-ttu-id="bfc6f-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-639">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-640">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="bfc6f-640">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="bfc6f-641">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-641">- Content</span></span><br><span data-ttu-id="bfc6f-642">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-642">
         - TaskPane</span></span><br><span data-ttu-id="bfc6f-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-643">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-644">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-645">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="bfc6f-645">- ActiveView</span></span><br><span data-ttu-id="bfc6f-646">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-646">
         - CompressedFile</span></span><br><span data-ttu-id="bfc6f-647">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-647">
         - DocumentEvents</span></span><br><span data-ttu-id="bfc6f-648">
         - File</span><span class="sxs-lookup"><span data-stu-id="bfc6f-648">
         - File</span></span><br><span data-ttu-id="bfc6f-649">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-649">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="bfc6f-650">
         - PdfFile</span></span><br><span data-ttu-id="bfc6f-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-651">
         - Selection</span></span><br><span data-ttu-id="bfc6f-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-652">
         - Settings</span></span><br><span data-ttu-id="bfc6f-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-653">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="onenote"></a><span data-ttu-id="bfc6f-654">OneNote</span><span class="sxs-lookup"><span data-stu-id="bfc6f-654">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bfc6f-655">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-655">Platform</span></span></th>
    <th><span data-ttu-id="bfc6f-656">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc6f-656">Extension points</span></span></th>
    <th><span data-ttu-id="bfc6f-657">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-657">API requirement sets</span></span></th>
    <th><span data-ttu-id="bfc6f-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-658"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-659">Office Online</span><span class="sxs-lookup"><span data-stu-id="bfc6f-659">Office Online</span></span></td>
    <td> <span data-ttu-id="bfc6f-660">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-660">- Content</span></span><br><span data-ttu-id="bfc6f-661">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-661">
         - TaskPane</span></span><br><span data-ttu-id="bfc6f-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-662">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-663">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="bfc6f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-665">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="bfc6f-665">- DocumentEvents</span></span><br><span data-ttu-id="bfc6f-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="bfc6f-667">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-667">
         - ImageCoercion</span></span><br><span data-ttu-id="bfc6f-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="bfc6f-668">
         - Settings</span></span><br><span data-ttu-id="bfc6f-669">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-669">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="bfc6f-670">Project</span><span class="sxs-lookup"><span data-stu-id="bfc6f-670">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="bfc6f-671">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="bfc6f-671">Platform</span></span></th>
    <th><span data-ttu-id="bfc6f-672">拡張点</span><span class="sxs-lookup"><span data-stu-id="bfc6f-672">Extension points</span></span></th>
    <th><span data-ttu-id="bfc6f-673">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-673">API requirement sets</span></span></th>
    <th><span data-ttu-id="bfc6f-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-674"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-675">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-675">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-676">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-676">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bfc6f-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-677">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-678">- Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-678">- Selection</span></span><br><span data-ttu-id="bfc6f-679">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-679">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-680">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-680">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-681">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-681">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bfc6f-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-682">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-683">- Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-683">- Selection</span></span><br><span data-ttu-id="bfc6f-684">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-684">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="bfc6f-685">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="bfc6f-685">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="bfc6f-686">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bfc6f-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="bfc6f-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="bfc6f-687">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="bfc6f-688">- Selection</span><span class="sxs-lookup"><span data-stu-id="bfc6f-688">- Selection</span></span><br><span data-ttu-id="bfc6f-689">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="bfc6f-689">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="bfc6f-690">関連項目</span><span class="sxs-lookup"><span data-stu-id="bfc6f-690">See also</span></span>

- [<span data-ttu-id="bfc6f-691">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="bfc6f-691">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="bfc6f-692">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-692">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="bfc6f-693">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="bfc6f-693">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="bfc6f-694">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="bfc6f-694">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
