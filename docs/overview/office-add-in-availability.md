---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: fe5b1d1278d2c14192fb6fd212f24bb08571d35d
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691126"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5da51-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5da51-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="5da51-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="5da51-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5da51-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="5da51-108">パッケージ版 Office 2019 のビルド番号は 16.0.10827.20150 です。</span><span class="sxs-lookup"><span data-stu-id="5da51-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="5da51-109">Excel</span><span class="sxs-lookup"><span data-stu-id="5da51-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5da51-110">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5da51-111">拡張点</span><span class="sxs-lookup"><span data-stu-id="5da51-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5da51-112">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5da51-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5da51-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="5da51-114">Office Online</span></span></td>
    <td> <span data-ttu-id="5da51-115">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-115">- TaskPane</span></span><br><span data-ttu-id="5da51-116">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-116">
        - Content</span></span><br><span data-ttu-id="5da51-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="5da51-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5da51-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5da51-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5da51-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5da51-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5da51-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5da51-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5da51-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5da51-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5da51-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5da51-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-127">
        - BindingEvents</span></span><br><span data-ttu-id="5da51-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-128">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-129">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-130">
        - File</span></span><br><span data-ttu-id="5da51-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-131">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-133">
        - Selection</span></span><br><span data-ttu-id="5da51-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-134">
        - Settings</span></span><br><span data-ttu-id="5da51-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-135">
        - TableBindings</span></span><br><span data-ttu-id="5da51-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-136">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-137">
        - TextBindings</span></span><br><span data-ttu-id="5da51-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-139">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-140">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-140">- TaskPane</span></span><br><span data-ttu-id="5da51-141">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-141">
        - Content</span></span><br><span data-ttu-id="5da51-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="5da51-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5da51-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5da51-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5da51-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5da51-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5da51-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5da51-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5da51-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5da51-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5da51-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5da51-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-152">
        - BindingEvents</span></span><br><span data-ttu-id="5da51-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-153">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-154">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-155">
        - File</span></span><br><span data-ttu-id="5da51-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-156">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-158">
        - Selection</span></span><br><span data-ttu-id="5da51-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-159">
        - Settings</span></span><br><span data-ttu-id="5da51-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-160">
        - TableBindings</span></span><br><span data-ttu-id="5da51-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-161">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-162">
        - TextBindings</span></span><br><span data-ttu-id="5da51-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-164">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="5da51-165">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-165">- TaskPane</span></span><br><span data-ttu-id="5da51-166">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-166">
        - Content</span></span><br><span data-ttu-id="5da51-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5da51-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5da51-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5da51-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5da51-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5da51-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5da51-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5da51-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5da51-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5da51-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5da51-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-177">- BindingEvents</span></span><br><span data-ttu-id="5da51-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-178">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-179">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-180">
        - File</span></span><br><span data-ttu-id="5da51-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-181">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-182">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-184">
        - Selection</span></span><br><span data-ttu-id="5da51-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-185">
        - Settings</span></span><br><span data-ttu-id="5da51-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-186">
        - TableBindings</span></span><br><span data-ttu-id="5da51-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-187">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-188">
        - TextBindings</span></span><br><span data-ttu-id="5da51-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-190">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="5da51-191">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-191">- TaskPane</span></span><br><span data-ttu-id="5da51-192">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-192">
        - Content</span></span></td>
    <td><span data-ttu-id="5da51-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="5da51-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-195">- BindingEvents</span></span><br><span data-ttu-id="5da51-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-196">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-197">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-198">
        - File</span></span><br><span data-ttu-id="5da51-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-199">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-200">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-202">
        - Selection</span></span><br><span data-ttu-id="5da51-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-203">
        - Settings</span></span><br><span data-ttu-id="5da51-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-204">
        - TableBindings</span></span><br><span data-ttu-id="5da51-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-205">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-206">
        - TextBindings</span></span><br><span data-ttu-id="5da51-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-208">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="5da51-209">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-209">
        - TaskPane</span></span><br><span data-ttu-id="5da51-210">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5da51-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5da51-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="5da51-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-212">
        - BindingEvents</span></span><br><span data-ttu-id="5da51-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-213">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-214">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-215">
        - File</span></span><br><span data-ttu-id="5da51-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-216">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-217">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-219">
        - Selection</span></span><br><span data-ttu-id="5da51-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-220">
        - Settings</span></span><br><span data-ttu-id="5da51-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-221">
        - TableBindings</span></span><br><span data-ttu-id="5da51-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-222">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-223">
        - TextBindings</span></span><br><span data-ttu-id="5da51-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-225">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="5da51-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="5da51-226">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-226">- TaskPane</span></span><br><span data-ttu-id="5da51-227">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-227">
        - Content</span></span></td>
    <td><span data-ttu-id="5da51-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5da51-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5da51-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5da51-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5da51-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5da51-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5da51-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5da51-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5da51-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5da51-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-237">- BindingEvents</span></span><br><span data-ttu-id="5da51-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-238">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-239">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-240">
        - File</span></span><br><span data-ttu-id="5da51-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-241">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-242">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-244">
        - Selection</span></span><br><span data-ttu-id="5da51-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-245">
        - Settings</span></span><br><span data-ttu-id="5da51-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-246">
        - TableBindings</span></span><br><span data-ttu-id="5da51-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-247">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-248">
        - TextBindings</span></span><br><span data-ttu-id="5da51-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-250">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="5da51-251">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-251">- TaskPane</span></span><br><span data-ttu-id="5da51-252">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-252">
        - Content</span></span><br><span data-ttu-id="5da51-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5da51-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5da51-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5da51-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5da51-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5da51-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5da51-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5da51-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5da51-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5da51-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5da51-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-263">- BindingEvents</span></span><br><span data-ttu-id="5da51-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-264">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-265">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-266">
        - File</span></span><br><span data-ttu-id="5da51-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-267">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-268">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-270">
        - PdfFile</span></span><br><span data-ttu-id="5da51-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-271">
        - Selection</span></span><br><span data-ttu-id="5da51-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-272">
        - Settings</span></span><br><span data-ttu-id="5da51-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-273">
        - TableBindings</span></span><br><span data-ttu-id="5da51-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-274">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-275">
        - TextBindings</span></span><br><span data-ttu-id="5da51-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-277">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="5da51-278">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-278">- TaskPane</span></span><br><span data-ttu-id="5da51-279">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-279">
        - Content</span></span><br><span data-ttu-id="5da51-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5da51-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5da51-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5da51-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5da51-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5da51-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5da51-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5da51-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5da51-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5da51-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5da51-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-290">- BindingEvents</span></span><br><span data-ttu-id="5da51-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-291">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-292">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-293">
        - File</span></span><br><span data-ttu-id="5da51-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-294">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-295">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-297">
        - PdfFile</span></span><br><span data-ttu-id="5da51-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-298">
        - Selection</span></span><br><span data-ttu-id="5da51-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-299">
        - Settings</span></span><br><span data-ttu-id="5da51-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-300">
        - TableBindings</span></span><br><span data-ttu-id="5da51-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-301">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-302">
        - TextBindings</span></span><br><span data-ttu-id="5da51-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-304">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="5da51-305">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-305">- TaskPane</span></span><br><span data-ttu-id="5da51-306">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-306">
        - Content</span></span></td>
    <td><span data-ttu-id="5da51-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5da51-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="5da51-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-309">- BindingEvents</span></span><br><span data-ttu-id="5da51-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-310">
        - CompressedFile</span></span><br><span data-ttu-id="5da51-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-311">
        - DocumentEvents</span></span><br><span data-ttu-id="5da51-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="5da51-312">
        - File</span></span><br><span data-ttu-id="5da51-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-313">
        - ImageCoercion</span></span><br><span data-ttu-id="5da51-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-314">
        - MatrixBindings</span></span><br><span data-ttu-id="5da51-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="5da51-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-316">
        - PdfFile</span></span><br><span data-ttu-id="5da51-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-317">
        - Selection</span></span><br><span data-ttu-id="5da51-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-318">
        - Settings</span></span><br><span data-ttu-id="5da51-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-319">
        - TableBindings</span></span><br><span data-ttu-id="5da51-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-320">
        - TableCoercion</span></span><br><span data-ttu-id="5da51-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-321">
        - TextBindings</span></span><br><span data-ttu-id="5da51-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5da51-323">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5da51-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="5da51-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="5da51-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5da51-325">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-325">Platform</span></span></th>
    <th><span data-ttu-id="5da51-326">拡張点</span><span class="sxs-lookup"><span data-stu-id="5da51-326">Extension points</span></span></th>
    <th><span data-ttu-id="5da51-327">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="5da51-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5da51-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="5da51-329">Office Online</span></span></td>
    <td> <span data-ttu-id="5da51-330">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-330">- Mail Read</span></span><br><span data-ttu-id="5da51-331">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-331">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5da51-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5da51-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5da51-340">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-341">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-342">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-342">- Mail Read</span></span><br><span data-ttu-id="5da51-343">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-343">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5da51-345">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5da51-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5da51-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5da51-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5da51-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5da51-353">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-354">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-355">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-355">- Mail Read</span></span><br><span data-ttu-id="5da51-356">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-356">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5da51-358">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5da51-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5da51-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5da51-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5da51-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5da51-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5da51-366">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-367">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-368">- Mail Read</span></span><br><span data-ttu-id="5da51-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-369">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5da51-371">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5da51-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5da51-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="5da51-376">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-377">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-378">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-378">- Mail Read</span></span><br><span data-ttu-id="5da51-379">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="5da51-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="5da51-384">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-385">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="5da51-385">Office 365 for iOS</span></span></td>
    <td> <span data-ttu-id="5da51-386">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-386">- Mail Read</span></span><br><span data-ttu-id="5da51-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5da51-393">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-394">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-395">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-395">- Mail Read</span></span><br><span data-ttu-id="5da51-396">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-396">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5da51-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5da51-404">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-405">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-406">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-406">- Mail Read</span></span><br><span data-ttu-id="5da51-407">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-407">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5da51-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5da51-415">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-416">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-417">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-417">- Mail Read</span></span><br><span data-ttu-id="5da51-418">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5da51-418">
      - Mail Compose</span></span><br><span data-ttu-id="5da51-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5da51-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5da51-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5da51-426">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-427">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="5da51-427">Office 365 for Android</span></span></td>
    <td> <span data-ttu-id="5da51-428">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5da51-428">- Mail Read</span></span><br><span data-ttu-id="5da51-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5da51-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5da51-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5da51-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5da51-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5da51-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5da51-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5da51-435">利用不可</span><span class="sxs-lookup"><span data-stu-id="5da51-435">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="5da51-436">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5da51-436">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="5da51-437">Word</span><span class="sxs-lookup"><span data-stu-id="5da51-437">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5da51-438">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-438">Platform</span></span></th>
    <th><span data-ttu-id="5da51-439">拡張点</span><span class="sxs-lookup"><span data-stu-id="5da51-439">Extension points</span></span></th>
    <th><span data-ttu-id="5da51-440">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-440">API requirement sets</span></span></th>
    <th><span data-ttu-id="5da51-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5da51-441"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-442">Office Online</span><span class="sxs-lookup"><span data-stu-id="5da51-442">Office Online</span></span></td>
    <td> <span data-ttu-id="5da51-443">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-443">- TaskPane</span></span><br><span data-ttu-id="5da51-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-444">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-445">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5da51-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5da51-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-448">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-449">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-449">- BindingEvents</span></span><br><span data-ttu-id="5da51-450">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-450">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-451">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-451">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-452">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-452">
         - File</span></span><br><span data-ttu-id="5da51-453">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-453">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-454">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-454">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-455">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-455">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-456">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-456">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-457">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-457">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-458">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-458">
         - PdfFile</span></span><br><span data-ttu-id="5da51-459">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-459">
         - Selection</span></span><br><span data-ttu-id="5da51-460">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-460">
         - Settings</span></span><br><span data-ttu-id="5da51-461">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-461">
         - TableBindings</span></span><br><span data-ttu-id="5da51-462">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-462">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-463">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-463">
         - TextBindings</span></span><br><span data-ttu-id="5da51-464">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-464">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-465">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-465">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-466">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-466">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-467">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-467">- TaskPane</span></span><br><span data-ttu-id="5da51-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-468">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-469">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5da51-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5da51-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-472">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-473">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-473">- BindingEvents</span></span><br><span data-ttu-id="5da51-474">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-474">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-475">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-475">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-476">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-476">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-477">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-477">
         - File</span></span><br><span data-ttu-id="5da51-478">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-478">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-479">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-479">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-480">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-480">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-481">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-481">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-482">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-482">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-483">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-483">
         - PdfFile</span></span><br><span data-ttu-id="5da51-484">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-484">
         - Selection</span></span><br><span data-ttu-id="5da51-485">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-485">
         - Settings</span></span><br><span data-ttu-id="5da51-486">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-486">
         - TableBindings</span></span><br><span data-ttu-id="5da51-487">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-487">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-488">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-488">
         - TextBindings</span></span><br><span data-ttu-id="5da51-489">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-489">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-490">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-490">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-491">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-491">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-492">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-492">- TaskPane</span></span><br><span data-ttu-id="5da51-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-493">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-494">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5da51-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5da51-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-497">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-498">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-498">- BindingEvents</span></span><br><span data-ttu-id="5da51-499">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-499">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-500">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-500">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-501">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-501">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-502">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-502">
         - File</span></span><br><span data-ttu-id="5da51-503">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-503">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-504">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-504">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-505">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-505">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-506">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-506">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-507">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-507">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-508">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-508">
         - PdfFile</span></span><br><span data-ttu-id="5da51-509">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-509">
         - Selection</span></span><br><span data-ttu-id="5da51-510">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-510">
         - Settings</span></span><br><span data-ttu-id="5da51-511">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-511">
         - TableBindings</span></span><br><span data-ttu-id="5da51-512">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-512">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-513">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-513">
         - TextBindings</span></span><br><span data-ttu-id="5da51-514">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-514">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-515">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-515">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-516">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-516">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-517">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-517">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-518">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-519">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="5da51-520">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-520">- BindingEvents</span></span><br><span data-ttu-id="5da51-521">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-521">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-522">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-522">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-523">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-523">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-524">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-524">
         - File</span></span><br><span data-ttu-id="5da51-525">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-525">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-526">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-526">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-527">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-527">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-528">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-528">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-529">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-529">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-530">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-530">
         - PdfFile</span></span><br><span data-ttu-id="5da51-531">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-531">
         - Selection</span></span><br><span data-ttu-id="5da51-532">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-532">
         - Settings</span></span><br><span data-ttu-id="5da51-533">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-533">
         - TableBindings</span></span><br><span data-ttu-id="5da51-534">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-534">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-535">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-535">
         - TextBindings</span></span><br><span data-ttu-id="5da51-536">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-536">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-537">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-537">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-538">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-538">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-539">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-539">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5da51-540">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5da51-541">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-541">- BindingEvents</span></span><br><span data-ttu-id="5da51-542">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-542">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-543">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-543">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-544">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-544">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-545">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-545">
         - File</span></span><br><span data-ttu-id="5da51-546">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-546">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-547">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-547">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-548">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-551">
         - PdfFile</span></span><br><span data-ttu-id="5da51-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-552">
         - Selection</span></span><br><span data-ttu-id="5da51-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-553">
         - Settings</span></span><br><span data-ttu-id="5da51-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-554">
         - TableBindings</span></span><br><span data-ttu-id="5da51-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-555">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-556">
         - TextBindings</span></span><br><span data-ttu-id="5da51-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-557">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-558">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-559">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="5da51-559">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="5da51-560">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-560">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-561">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5da51-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5da51-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5da51-564">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5da51-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-565">- BindingEvents</span></span><br><span data-ttu-id="5da51-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-566">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-568">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-569">
         - File</span></span><br><span data-ttu-id="5da51-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-571">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-571">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-572">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-572">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-573">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-573">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-574">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-574">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-575">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-575">
         - PdfFile</span></span><br><span data-ttu-id="5da51-576">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-576">
         - Selection</span></span><br><span data-ttu-id="5da51-577">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-577">
         - Settings</span></span><br><span data-ttu-id="5da51-578">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-578">
         - TableBindings</span></span><br><span data-ttu-id="5da51-579">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-579">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-580">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-580">
         - TextBindings</span></span><br><span data-ttu-id="5da51-581">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-581">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-582">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-582">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-583">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-583">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-584">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-584">- TaskPane</span></span><br><span data-ttu-id="5da51-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-585">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-586">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5da51-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5da51-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5da51-589">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5da51-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-590">- BindingEvents</span></span><br><span data-ttu-id="5da51-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-591">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-593">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-594">
         - File</span></span><br><span data-ttu-id="5da51-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-596">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-596">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-597">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-600">
         - PdfFile</span></span><br><span data-ttu-id="5da51-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-601">
         - Selection</span></span><br><span data-ttu-id="5da51-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-602">
         - Settings</span></span><br><span data-ttu-id="5da51-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-603">
         - TableBindings</span></span><br><span data-ttu-id="5da51-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-604">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-605">
         - TextBindings</span></span><br><span data-ttu-id="5da51-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-606">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-608">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-608">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-609">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-609">- TaskPane</span></span><br><span data-ttu-id="5da51-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-610">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-611">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5da51-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5da51-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5da51-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5da51-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5da51-614">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5da51-615">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-615">- BindingEvents</span></span><br><span data-ttu-id="5da51-616">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-616">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-617">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-617">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-618">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-618">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-619">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-619">
         - File</span></span><br><span data-ttu-id="5da51-620">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-620">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-621">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-621">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-622">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-622">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-623">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-623">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-624">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-624">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-625">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-625">
         - PdfFile</span></span><br><span data-ttu-id="5da51-626">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-626">
         - Selection</span></span><br><span data-ttu-id="5da51-627">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-627">
         - Settings</span></span><br><span data-ttu-id="5da51-628">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-628">
         - TableBindings</span></span><br><span data-ttu-id="5da51-629">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-629">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-630">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-630">
         - TextBindings</span></span><br><span data-ttu-id="5da51-631">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-631">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-632">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-632">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-633">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-633">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-634">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-634">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-635">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5da51-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5da51-636">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="5da51-637">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-637">- BindingEvents</span></span><br><span data-ttu-id="5da51-638">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-638">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-639">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5da51-639">
         - CustomXmlParts</span></span><br><span data-ttu-id="5da51-640">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-640">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-641">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-641">
         - File</span></span><br><span data-ttu-id="5da51-642">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-642">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-643">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-643">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-644">
         - MatrixBindings</span></span><br><span data-ttu-id="5da51-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="5da51-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5da51-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-647">
         - PdfFile</span></span><br><span data-ttu-id="5da51-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-648">
         - Selection</span></span><br><span data-ttu-id="5da51-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-649">
         - Settings</span></span><br><span data-ttu-id="5da51-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-650">
         - TableBindings</span></span><br><span data-ttu-id="5da51-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-651">
         - TableCoercion</span></span><br><span data-ttu-id="5da51-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5da51-652">
         - TextBindings</span></span><br><span data-ttu-id="5da51-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-653">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5da51-654">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="5da51-655">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5da51-655">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5da51-656">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5da51-656">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5da51-657">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-657">Platform</span></span></th>
    <th><span data-ttu-id="5da51-658">拡張点</span><span class="sxs-lookup"><span data-stu-id="5da51-658">Extension points</span></span></th>
    <th><span data-ttu-id="5da51-659">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-659">API requirement sets</span></span></th>
    <th><span data-ttu-id="5da51-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5da51-660"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-661">Office Online</span><span class="sxs-lookup"><span data-stu-id="5da51-661">Office Online</span></span></td>
    <td> <span data-ttu-id="5da51-662">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-662">- Content</span></span><br><span data-ttu-id="5da51-663">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-663">
         - TaskPane</span></span><br><span data-ttu-id="5da51-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-664">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-665">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-666">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-666">- ActiveView</span></span><br><span data-ttu-id="5da51-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-667">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-668">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-669">
         - File</span></span><br><span data-ttu-id="5da51-670">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-670">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-671">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-671">
         - PdfFile</span></span><br><span data-ttu-id="5da51-672">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-672">
         - Selection</span></span><br><span data-ttu-id="5da51-673">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-673">
         - Settings</span></span><br><span data-ttu-id="5da51-674">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-674">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-675">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-675">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-676">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-676">- Content</span></span><br><span data-ttu-id="5da51-677">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-677">
         - TaskPane</span></span><br><span data-ttu-id="5da51-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-678">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-679">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-680">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-680">- ActiveView</span></span><br><span data-ttu-id="5da51-681">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-681">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-682">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-682">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-683">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-683">
         - File</span></span><br><span data-ttu-id="5da51-684">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-684">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-685">
         - PdfFile</span></span><br><span data-ttu-id="5da51-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-686">
         - Selection</span></span><br><span data-ttu-id="5da51-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-687">
         - Settings</span></span><br><span data-ttu-id="5da51-688">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-688">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-689">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-689">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-690">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-690">- Content</span></span><br><span data-ttu-id="5da51-691">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-691">
         - TaskPane</span></span><br><span data-ttu-id="5da51-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-692">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-693">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-694">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-694">- ActiveView</span></span><br><span data-ttu-id="5da51-695">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-695">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-696">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-696">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-697">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-697">
         - File</span></span><br><span data-ttu-id="5da51-698">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-698">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-699">
         - PdfFile</span></span><br><span data-ttu-id="5da51-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-700">
         - Selection</span></span><br><span data-ttu-id="5da51-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-701">
         - Settings</span></span><br><span data-ttu-id="5da51-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-702">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-703">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-703">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-704">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-704">- Content</span></span><br><span data-ttu-id="5da51-705">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-705">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5da51-706">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5da51-707">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-707">- ActiveView</span></span><br><span data-ttu-id="5da51-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-708">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-709">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-710">
         - File</span></span><br><span data-ttu-id="5da51-711">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-711">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-712">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-712">
         - PdfFile</span></span><br><span data-ttu-id="5da51-713">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-713">
         - Selection</span></span><br><span data-ttu-id="5da51-714">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-714">
         - Settings</span></span><br><span data-ttu-id="5da51-715">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-715">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-716">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-716">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-717">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-717">- Content</span></span><br><span data-ttu-id="5da51-718">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-718">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="5da51-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5da51-719">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5da51-720">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-720">- ActiveView</span></span><br><span data-ttu-id="5da51-721">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-721">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-722">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-722">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-723">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-723">
         - File</span></span><br><span data-ttu-id="5da51-724">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-724">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-725">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-725">
         - PdfFile</span></span><br><span data-ttu-id="5da51-726">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-726">
         - Selection</span></span><br><span data-ttu-id="5da51-727">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-727">
         - Settings</span></span><br><span data-ttu-id="5da51-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-728">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-729">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="5da51-729">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="5da51-730">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-730">- Content</span></span><br><span data-ttu-id="5da51-731">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-731">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-732">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="5da51-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-733">- ActiveView</span></span><br><span data-ttu-id="5da51-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-734">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-735">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-736">
         - File</span></span><br><span data-ttu-id="5da51-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-737">
         - PdfFile</span></span><br><span data-ttu-id="5da51-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-738">
         - Selection</span></span><br><span data-ttu-id="5da51-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-739">
         - Settings</span></span><br><span data-ttu-id="5da51-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-740">
         - TextCoercion</span></span><br><span data-ttu-id="5da51-741">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-741">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-742">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-742">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-743">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-743">- Content</span></span><br><span data-ttu-id="5da51-744">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-744">
         - TaskPane</span></span><br><span data-ttu-id="5da51-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-745">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-746">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-747">- ActiveView</span></span><br><span data-ttu-id="5da51-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-748">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-749">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-750">
         - File</span></span><br><span data-ttu-id="5da51-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-751">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-752">
         - PdfFile</span></span><br><span data-ttu-id="5da51-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-753">
         - Selection</span></span><br><span data-ttu-id="5da51-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-754">
         - Settings</span></span><br><span data-ttu-id="5da51-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-756">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-756">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-757">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-757">- Content</span></span><br><span data-ttu-id="5da51-758">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-758">
         - TaskPane</span></span><br><span data-ttu-id="5da51-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-759">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-760">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-761">- ActiveView</span></span><br><span data-ttu-id="5da51-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-762">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-763">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-764">
         - File</span></span><br><span data-ttu-id="5da51-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-765">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-766">
         - PdfFile</span></span><br><span data-ttu-id="5da51-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-767">
         - Selection</span></span><br><span data-ttu-id="5da51-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-768">
         - Settings</span></span><br><span data-ttu-id="5da51-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-770">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5da51-770">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="5da51-771">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-771">- Content</span></span><br><span data-ttu-id="5da51-772">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-772">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5da51-773">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5da51-774">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5da51-774">- ActiveView</span></span><br><span data-ttu-id="5da51-775">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5da51-775">
         - CompressedFile</span></span><br><span data-ttu-id="5da51-776">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-776">
         - DocumentEvents</span></span><br><span data-ttu-id="5da51-777">
         - File</span><span class="sxs-lookup"><span data-stu-id="5da51-777">
         - File</span></span><br><span data-ttu-id="5da51-778">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-778">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-779">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5da51-779">
         - PdfFile</span></span><br><span data-ttu-id="5da51-780">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-780">
         - Selection</span></span><br><span data-ttu-id="5da51-781">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-781">
         - Settings</span></span><br><span data-ttu-id="5da51-782">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-782">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5da51-783">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5da51-783">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="5da51-784">OneNote</span><span class="sxs-lookup"><span data-stu-id="5da51-784">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5da51-785">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-785">Platform</span></span></th>
    <th><span data-ttu-id="5da51-786">拡張点</span><span class="sxs-lookup"><span data-stu-id="5da51-786">Extension points</span></span></th>
    <th><span data-ttu-id="5da51-787">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-787">API requirement sets</span></span></th>
    <th><span data-ttu-id="5da51-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5da51-788"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-789">Office Online</span><span class="sxs-lookup"><span data-stu-id="5da51-789">Office Online</span></span></td>
    <td> <span data-ttu-id="5da51-790">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5da51-790">- Content</span></span><br><span data-ttu-id="5da51-791">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-791">
         - TaskPane</span></span><br><span data-ttu-id="5da51-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5da51-792">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5da51-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-793">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5da51-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-794">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-795">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5da51-795">- DocumentEvents</span></span><br><span data-ttu-id="5da51-796">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-796">
         - HtmlCoercion</span></span><br><span data-ttu-id="5da51-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-797">
         - ImageCoercion</span></span><br><span data-ttu-id="5da51-798">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5da51-798">
         - Settings</span></span><br><span data-ttu-id="5da51-799">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-799">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="5da51-800">Project</span><span class="sxs-lookup"><span data-stu-id="5da51-800">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5da51-801">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5da51-801">Platform</span></span></th>
    <th><span data-ttu-id="5da51-802">拡張点</span><span class="sxs-lookup"><span data-stu-id="5da51-802">Extension points</span></span></th>
    <th><span data-ttu-id="5da51-803">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-803">API requirement sets</span></span></th>
    <th><span data-ttu-id="5da51-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5da51-804"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-805">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-805">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-806">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-806">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-807">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-808">- Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-808">- Selection</span></span><br><span data-ttu-id="5da51-809">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-809">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-810">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-810">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-811">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-811">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-812">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-813">- Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-813">- Selection</span></span><br><span data-ttu-id="5da51-814">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-814">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5da51-815">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="5da51-815">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="5da51-816">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5da51-816">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5da51-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5da51-817">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5da51-818">- Selection</span><span class="sxs-lookup"><span data-stu-id="5da51-818">- Selection</span></span><br><span data-ttu-id="5da51-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5da51-819">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5da51-820">関連項目</span><span class="sxs-lookup"><span data-stu-id="5da51-820">See also</span></span>

- [<span data-ttu-id="5da51-821">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="5da51-821">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5da51-822">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-822">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="5da51-823">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="5da51-823">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="5da51-824">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="5da51-824">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
