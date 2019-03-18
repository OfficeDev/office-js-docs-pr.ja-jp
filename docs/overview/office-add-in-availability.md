---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 03/15/2019
localization_priority: Priority
ms.openlocfilehash: 4348881c35e4c79975d34406e4668b2693405134
ms.sourcegitcommit: c4d6ecdc41ea67291b6d155c3b246e31ec2e38b7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/16/2019
ms.locfileid: "30654964"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4d9d9-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4d9d9-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d9d9-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4d9d9-p102">MSI からインストールされた Office 2016 のビルド番号は、16.0.4266.1001 です。このバージョンには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d9d9-p102">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span>
>
> <span data-ttu-id="4d9d9-108">パッケージ版 Office 2019 のビルド番号は 16.0.10827.20150 です。</span><span class="sxs-lookup"><span data-stu-id="4d9d9-108">The build number for a one-time purchase of Office 2019 is 16.0.10827.20150.</span></span>

## <a name="excel"></a><span data-ttu-id="4d9d9-109">Excel</span><span class="sxs-lookup"><span data-stu-id="4d9d9-109">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4d9d9-110">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-110">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4d9d9-111">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d9d9-111">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4d9d9-112">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-112">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4d9d9-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-113"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d9d9-114">Office Online</span></span></td>
    <td> <span data-ttu-id="4d9d9-115">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-115">- TaskPane</span></span><br><span data-ttu-id="4d9d9-116">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-116">
        - Content</span></span><br><span data-ttu-id="4d9d9-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4d9d9-117">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4d9d9-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-118">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-119">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-120">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-121">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d9d9-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-122">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d9d9-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-123">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d9d9-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-124">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d9d9-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-125">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d9d9-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-126">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d9d9-127">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-127">
        - BindingEvents</span></span><br><span data-ttu-id="4d9d9-128">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-128">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-129">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-129">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-130">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-130">
        - File</span></span><br><span data-ttu-id="4d9d9-131">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-131">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-132">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-132">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-133">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-133">
        - Selection</span></span><br><span data-ttu-id="4d9d9-134">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-134">
        - Settings</span></span><br><span data-ttu-id="4d9d9-135">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-135">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-136">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-136">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-137">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-137">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-138">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-138">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-139">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-139">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-140">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-140">- TaskPane</span></span><br><span data-ttu-id="4d9d9-141">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-141">
        - Content</span></span><br><span data-ttu-id="4d9d9-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4d9d9-142">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4d9d9-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-143">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-144">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-145">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-146">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d9d9-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-147">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d9d9-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-148">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d9d9-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-149">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d9d9-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-150">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d9d9-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-151">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d9d9-152">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-152">
        - BindingEvents</span></span><br><span data-ttu-id="4d9d9-153">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-153">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-154">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-154">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-155">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-155">
        - File</span></span><br><span data-ttu-id="4d9d9-156">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-156">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-157">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-157">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-158">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-158">
        - Selection</span></span><br><span data-ttu-id="4d9d9-159">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-159">
        - Settings</span></span><br><span data-ttu-id="4d9d9-160">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-160">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-161">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-161">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-162">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-162">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-163">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-163">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-164">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-164">Office 2019 for Windows</span></span></td>
    <td><span data-ttu-id="4d9d9-165">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-165">- TaskPane</span></span><br><span data-ttu-id="4d9d9-166">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-166">
        - Content</span></span><br><span data-ttu-id="4d9d9-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-167">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d9d9-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-168">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-169">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-170">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-171">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d9d9-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-172">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d9d9-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-173">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d9d9-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-174">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d9d9-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-175">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d9d9-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-176">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d9d9-177">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-177">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-178">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-178">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-179">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-179">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-180">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-180">
        - File</span></span><br><span data-ttu-id="4d9d9-181">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-181">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-182">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-182">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-183">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-183">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-184">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-184">
        - Selection</span></span><br><span data-ttu-id="4d9d9-185">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-185">
        - Settings</span></span><br><span data-ttu-id="4d9d9-186">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-186">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-187">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-187">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-188">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-188">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-189">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-189">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-190">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-190">Office 2016 for Windows</span></span></td>
    <td><span data-ttu-id="4d9d9-191">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-191">- TaskPane</span></span><br><span data-ttu-id="4d9d9-192">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-192">
        - Content</span></span></td>
    <td><span data-ttu-id="4d9d9-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-193">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-194">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4d9d9-195">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-195">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-196">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-196">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-197">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-197">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-198">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-198">
        - File</span></span><br><span data-ttu-id="4d9d9-199">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-199">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-200">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-200">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-201">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-201">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-202">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-202">
        - Selection</span></span><br><span data-ttu-id="4d9d9-203">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-203">
        - Settings</span></span><br><span data-ttu-id="4d9d9-204">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-204">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-205">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-205">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-206">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-206">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-207">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-207">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-208">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-208">Office 2013 for Windows</span></span></td>
    <td><span data-ttu-id="4d9d9-209">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-209">
        - TaskPane</span></span><br><span data-ttu-id="4d9d9-210">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-210">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4d9d9-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-211">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="4d9d9-212">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-212">
        - BindingEvents</span></span><br><span data-ttu-id="4d9d9-213">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-213">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-214">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-214">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-215">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-215">
        - File</span></span><br><span data-ttu-id="4d9d9-216">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-216">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-217">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-217">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-218">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-218">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-219">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-219">
        - Selection</span></span><br><span data-ttu-id="4d9d9-220">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-220">
        - Settings</span></span><br><span data-ttu-id="4d9d9-221">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-221">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-222">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-222">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-223">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-223">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-224">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-224">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-225">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="4d9d9-225">Office 365 for iPad</span></span></td>
    <td><span data-ttu-id="4d9d9-226">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-226">- TaskPane</span></span><br><span data-ttu-id="4d9d9-227">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-227">
        - Content</span></span></td>
    <td><span data-ttu-id="4d9d9-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-228">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-229">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-230">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-231">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d9d9-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-232">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d9d9-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-233">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d9d9-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-234">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d9d9-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-235">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d9d9-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-236">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d9d9-237">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-237">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-238">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-238">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-239">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-239">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-240">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-240">
        - File</span></span><br><span data-ttu-id="4d9d9-241">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-241">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-242">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-242">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-243">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-243">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-244">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-244">
        - Selection</span></span><br><span data-ttu-id="4d9d9-245">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-245">
        - Settings</span></span><br><span data-ttu-id="4d9d9-246">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-246">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-247">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-247">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-248">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-248">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-249">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-249">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-250">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-250">Office 365 for Mac</span></span></td>
    <td><span data-ttu-id="4d9d9-251">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-251">- TaskPane</span></span><br><span data-ttu-id="4d9d9-252">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-252">
        - Content</span></span><br><span data-ttu-id="4d9d9-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-253">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d9d9-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-254">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-255">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-256">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-257">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d9d9-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-258">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d9d9-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-259">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d9d9-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-260">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d9d9-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-261">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d9d9-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-262">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d9d9-263">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-263">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-264">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-264">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-265">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-265">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-266">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-266">
        - File</span></span><br><span data-ttu-id="4d9d9-267">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-267">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-268">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-268">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-269">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-269">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-270">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-270">
        - PdfFile</span></span><br><span data-ttu-id="4d9d9-271">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-271">
        - Selection</span></span><br><span data-ttu-id="4d9d9-272">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-272">
        - Settings</span></span><br><span data-ttu-id="4d9d9-273">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-273">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-274">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-274">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-275">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-275">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-276">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-276">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-277">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-277">Office 2019 for Mac</span></span></td>
    <td><span data-ttu-id="4d9d9-278">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-278">- TaskPane</span></span><br><span data-ttu-id="4d9d9-279">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-279">
        - Content</span></span><br><span data-ttu-id="4d9d9-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-280">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d9d9-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-281">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-282">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-283">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-284">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d9d9-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-285">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d9d9-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-286">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d9d9-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-287">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d9d9-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-288">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d9d9-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-289">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d9d9-290">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-290">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-291">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-291">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-292">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-292">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-293">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-293">
        - File</span></span><br><span data-ttu-id="4d9d9-294">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-294">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-295">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-295">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-296">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-296">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-297">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-297">
        - PdfFile</span></span><br><span data-ttu-id="4d9d9-298">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-298">
        - Selection</span></span><br><span data-ttu-id="4d9d9-299">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-299">
        - Settings</span></span><br><span data-ttu-id="4d9d9-300">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-300">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-301">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-301">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-302">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-302">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-303">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-303">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-304">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-304">Office 2016 for Mac</span></span></td>
    <td><span data-ttu-id="4d9d9-305">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-305">- TaskPane</span></span><br><span data-ttu-id="4d9d9-306">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-306">
        - Content</span></span></td>
    <td><span data-ttu-id="4d9d9-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-307">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-308">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="4d9d9-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-309">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-310">
        - CompressedFile</span></span><br><span data-ttu-id="4d9d9-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-311">
        - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-312">
        - File</span></span><br><span data-ttu-id="4d9d9-313">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-313">
        - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-314">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-314">
        - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-315">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-315">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-316">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-316">
        - PdfFile</span></span><br><span data-ttu-id="4d9d9-317">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-317">
        - Selection</span></span><br><span data-ttu-id="4d9d9-318">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-318">
        - Settings</span></span><br><span data-ttu-id="4d9d9-319">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-319">
        - TableBindings</span></span><br><span data-ttu-id="4d9d9-320">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-320">
        - TableCoercion</span></span><br><span data-ttu-id="4d9d9-321">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-321">
        - TextBindings</span></span><br><span data-ttu-id="4d9d9-322">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-322">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d9d9-323">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-323">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="outlook"></a><span data-ttu-id="4d9d9-324">Outlook</span><span class="sxs-lookup"><span data-stu-id="4d9d9-324">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d9d9-325">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-325">Platform</span></span></th>
    <th><span data-ttu-id="4d9d9-326">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d9d9-326">Extension points</span></span></th>
    <th><span data-ttu-id="4d9d9-327">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-327">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d9d9-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-328"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-329">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d9d9-329">Office Online</span></span></td>
    <td> <span data-ttu-id="4d9d9-330">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-330">- Mail Read</span></span><br><span data-ttu-id="4d9d9-331">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-331">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-332">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-333">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-334">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-335">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-336">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-337">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d9d9-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-338">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d9d9-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-339">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d9d9-340">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-340">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-341">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-341">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-342">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-342">- Mail Read</span></span><br><span data-ttu-id="4d9d9-343">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-343">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-344">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d9d9-345">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4d9d9-345">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d9d9-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-346">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-347">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-348">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-349">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-350">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d9d9-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-351">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d9d9-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-352">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d9d9-353">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-353">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-354">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-354">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-355">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-355">- Mail Read</span></span><br><span data-ttu-id="4d9d9-356">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-356">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-357">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d9d9-358">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4d9d9-358">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d9d9-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-359">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-360">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-361">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-362">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-363">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d9d9-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-364">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d9d9-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-365">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d9d9-366">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-366">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-367">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-367">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-368">- Mail Read</span></span><br><span data-ttu-id="4d9d9-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-369">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-370">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d9d9-371">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4d9d9-371">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d9d9-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-372">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-373">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-374">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-375">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4d9d9-376">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-377">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-377">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-378">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-378">- Mail Read</span></span><br><span data-ttu-id="4d9d9-379">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-379">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="4d9d9-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-380">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-381">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-382">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-383">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span></td>
    <td><span data-ttu-id="4d9d9-384">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-385">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="4d9d9-385">Office for iOS</span></span></td>
    <td> <span data-ttu-id="4d9d9-386">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-386">- Mail Read</span></span><br><span data-ttu-id="4d9d9-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-387">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-388">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-389">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-390">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-391">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-392">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4d9d9-393">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-393">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-394">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-394">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-395">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-395">- Mail Read</span></span><br><span data-ttu-id="4d9d9-396">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-396">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-397">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-398">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-399">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-400">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-401">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-402">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d9d9-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-403">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d9d9-404">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-404">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-405">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-405">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-406">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-406">- Mail Read</span></span><br><span data-ttu-id="4d9d9-407">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-407">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-408">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-409">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-410">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-411">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-412">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-413">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d9d9-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-414">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d9d9-415">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-416">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-416">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-417">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-417">- Mail Read</span></span><br><span data-ttu-id="4d9d9-418">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d9d9-418">
      - Mail Compose</span></span><br><span data-ttu-id="4d9d9-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-419">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-420">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-421">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-422">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-423">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-424">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d9d9-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-425">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d9d9-426">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-427">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="4d9d9-427">Office 365 SDK for Android</span></span></td>
    <td> <span data-ttu-id="4d9d9-428">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d9d9-428">- Mail Read</span></span><br><span data-ttu-id="4d9d9-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-429">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-430">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d9d9-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-431">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d9d9-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-432">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d9d9-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-433">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d9d9-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-434">
      - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4d9d9-435">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d9d9-435">Not available</span></span></td>
  </tr>
</table>

<br/>

## <a name="word"></a><span data-ttu-id="4d9d9-436">Word</span><span class="sxs-lookup"><span data-stu-id="4d9d9-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d9d9-437">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-437">Platform</span></span></th>
    <th><span data-ttu-id="4d9d9-438">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d9d9-438">Extension points</span></span></th>
    <th><span data-ttu-id="4d9d9-439">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-439">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d9d9-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-440"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d9d9-441">Office Online</span></span></td>
    <td> <span data-ttu-id="4d9d9-442">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-442">- TaskPane</span></span><br><span data-ttu-id="4d9d9-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-443">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-444">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-445">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-446">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-447">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-448">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-448">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-449">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-449">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-450">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-450">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-451">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-451">
         - File</span></span><br><span data-ttu-id="4d9d9-452">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-452">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-453">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-453">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-454">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-454">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-455">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-455">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-456">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-456">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-457">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-457">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-458">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-458">
         - Selection</span></span><br><span data-ttu-id="4d9d9-459">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-459">
         - Settings</span></span><br><span data-ttu-id="4d9d9-460">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-460">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-461">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-461">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-462">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-462">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-463">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-463">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-464">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-464">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-465">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-465">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-466">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-466">- TaskPane</span></span><br><span data-ttu-id="4d9d9-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-467">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-468">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-469">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-470">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-471">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-472">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-472">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-473">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-473">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-474">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-474">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-475">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-475">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-476">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-476">
         - File</span></span><br><span data-ttu-id="4d9d9-477">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-477">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-478">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-478">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-479">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-479">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-480">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-480">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-481">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-481">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-482">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-482">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-483">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-483">
         - Selection</span></span><br><span data-ttu-id="4d9d9-484">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-484">
         - Settings</span></span><br><span data-ttu-id="4d9d9-485">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-485">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-486">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-486">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-487">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-487">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-488">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-488">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-489">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-489">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-490">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-490">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-491">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-491">- TaskPane</span></span><br><span data-ttu-id="4d9d9-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-492">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-493">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-494">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-495">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-496">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-497">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-497">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-498">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-498">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-499">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-499">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-500">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-500">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-501">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-501">
         - File</span></span><br><span data-ttu-id="4d9d9-502">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-502">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-503">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-503">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-504">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-504">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-505">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-505">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-506">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-506">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-507">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-507">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-508">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-508">
         - Selection</span></span><br><span data-ttu-id="4d9d9-509">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-509">
         - Settings</span></span><br><span data-ttu-id="4d9d9-510">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-510">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-511">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-511">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-512">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-512">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-513">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-513">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-514">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-514">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-515">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-515">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-516">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-516">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-517">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-518">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4d9d9-519">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-519">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-520">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-520">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-521">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-521">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-522">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-522">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-523">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-523">
         - File</span></span><br><span data-ttu-id="4d9d9-524">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-524">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-525">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-525">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-526">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-526">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-527">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-527">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-528">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-528">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-529">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-529">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-530">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-530">
         - Selection</span></span><br><span data-ttu-id="4d9d9-531">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-531">
         - Settings</span></span><br><span data-ttu-id="4d9d9-532">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-532">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-533">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-533">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-534">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-534">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-535">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-535">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-536">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-536">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-537">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-537">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-538">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-538">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-539">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d9d9-540">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-540">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-541">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-541">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-542">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-542">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-543">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-543">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-544">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-544">
         - File</span></span><br><span data-ttu-id="4d9d9-545">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-545">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-546">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-546">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-547">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-547">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-548">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-548">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-549">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-549">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-550">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-550">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-551">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-551">
         - Selection</span></span><br><span data-ttu-id="4d9d9-552">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-552">
         - Settings</span></span><br><span data-ttu-id="4d9d9-553">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-553">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-554">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-554">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-555">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-555">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-556">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-556">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-557">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-557">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-558">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="4d9d9-558">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4d9d9-559">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-559">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-560">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-561">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-562">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4d9d9-563">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4d9d9-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-564">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-565">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-567">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-568">
         - File</span></span><br><span data-ttu-id="4d9d9-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-570">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-570">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-571">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-574">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-575">
         - Selection</span></span><br><span data-ttu-id="4d9d9-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-576">
         - Settings</span></span><br><span data-ttu-id="4d9d9-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-577">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-578">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-579">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-580">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-582">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-582">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-583">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-583">- TaskPane</span></span><br><span data-ttu-id="4d9d9-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-584">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-585">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-586">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-587">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4d9d9-588">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4d9d9-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-589">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-590">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-592">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-593">
         - File</span></span><br><span data-ttu-id="4d9d9-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-595">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-596">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-599">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-600">
         - Selection</span></span><br><span data-ttu-id="4d9d9-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-601">
         - Settings</span></span><br><span data-ttu-id="4d9d9-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-602">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-603">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-604">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-605">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-607">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-607">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-608">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-608">- TaskPane</span></span><br><span data-ttu-id="4d9d9-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-609">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-610">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-611">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="4d9d9-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-612">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="4d9d9-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="4d9d9-613">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="4d9d9-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-614">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-615">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-617">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-618">
         - File</span></span><br><span data-ttu-id="4d9d9-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-620">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-621">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-624">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-625">
         - Selection</span></span><br><span data-ttu-id="4d9d9-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-626">
         - Settings</span></span><br><span data-ttu-id="4d9d9-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-627">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-628">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-629">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-630">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-632">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-632">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-633">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-634">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-635">
        - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="4d9d9-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-636">- BindingEvents</span></span><br><span data-ttu-id="4d9d9-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-637">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d9d9-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d9d9-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-639">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-640">
         - File</span></span><br><span data-ttu-id="4d9d9-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-642">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-642">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-643">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-643">
         - MatrixBindings</span></span><br><span data-ttu-id="4d9d9-644">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-644">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d9d9-645">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-645">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d9d9-646">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-646">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-647">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-647">
         - Selection</span></span><br><span data-ttu-id="4d9d9-648">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-648">
         - Settings</span></span><br><span data-ttu-id="4d9d9-649">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-649">
         - TableBindings</span></span><br><span data-ttu-id="4d9d9-650">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-650">
         - TableCoercion</span></span><br><span data-ttu-id="4d9d9-651">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-651">
         - TextBindings</span></span><br><span data-ttu-id="4d9d9-652">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-652">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-653">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-653">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="4d9d9-654">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-654">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4d9d9-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4d9d9-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d9d9-656">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-656">Platform</span></span></th>
    <th><span data-ttu-id="4d9d9-657">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d9d9-657">Extension points</span></span></th>
    <th><span data-ttu-id="4d9d9-658">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-658">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d9d9-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-659"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d9d9-660">Office Online</span></span></td>
    <td> <span data-ttu-id="4d9d9-661">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-661">- Content</span></span><br><span data-ttu-id="4d9d9-662">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-662">
         - TaskPane</span></span><br><span data-ttu-id="4d9d9-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-663">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-664">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-665">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-665">- ActiveView</span></span><br><span data-ttu-id="4d9d9-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-666">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-667">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-667">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-668">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-668">
         - File</span></span><br><span data-ttu-id="4d9d9-669">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-669">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-670">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-671">
         - Selection</span></span><br><span data-ttu-id="4d9d9-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-672">
         - Settings</span></span><br><span data-ttu-id="4d9d9-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-673">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-674">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-674">Office 365 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-675">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-675">- Content</span></span><br><span data-ttu-id="4d9d9-676">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-676">
         - TaskPane</span></span><br><span data-ttu-id="4d9d9-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-677">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-678">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-679">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-679">- ActiveView</span></span><br><span data-ttu-id="4d9d9-680">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-680">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-681">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-681">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-682">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-682">
         - File</span></span><br><span data-ttu-id="4d9d9-683">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-683">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-684">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-684">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-685">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-685">
         - Selection</span></span><br><span data-ttu-id="4d9d9-686">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-686">
         - Settings</span></span><br><span data-ttu-id="4d9d9-687">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-687">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-688">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-688">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-689">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-689">- Content</span></span><br><span data-ttu-id="4d9d9-690">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-690">
         - TaskPane</span></span><br><span data-ttu-id="4d9d9-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-691">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-692">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-693">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-693">- ActiveView</span></span><br><span data-ttu-id="4d9d9-694">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-694">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-695">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-695">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-696">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-696">
         - File</span></span><br><span data-ttu-id="4d9d9-697">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-697">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-698">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-699">
         - Selection</span></span><br><span data-ttu-id="4d9d9-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-700">
         - Settings</span></span><br><span data-ttu-id="4d9d9-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-701">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-702">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-702">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-703">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-703">- Content</span></span><br><span data-ttu-id="4d9d9-704">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-704">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-705">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d9d9-706">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-706">- ActiveView</span></span><br><span data-ttu-id="4d9d9-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-707">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-708">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-709">
         - File</span></span><br><span data-ttu-id="4d9d9-710">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-710">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-711">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-711">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-712">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-712">
         - Selection</span></span><br><span data-ttu-id="4d9d9-713">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-713">
         - Settings</span></span><br><span data-ttu-id="4d9d9-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-714">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-715">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-715">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-716">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-716">- Content</span></span><br><span data-ttu-id="4d9d9-717">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-717">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4d9d9-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-718">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d9d9-719">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-719">- ActiveView</span></span><br><span data-ttu-id="4d9d9-720">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-720">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-721">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-721">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-722">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-722">
         - File</span></span><br><span data-ttu-id="4d9d9-723">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-723">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-724">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-724">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-725">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-725">
         - Selection</span></span><br><span data-ttu-id="4d9d9-726">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-726">
         - Settings</span></span><br><span data-ttu-id="4d9d9-727">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-727">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-728">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="4d9d9-728">Office 365 for iPad</span></span></td>
    <td> <span data-ttu-id="4d9d9-729">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-729">- Content</span></span><br><span data-ttu-id="4d9d9-730">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-730">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-731">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
     <td> <span data-ttu-id="4d9d9-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-732">- ActiveView</span></span><br><span data-ttu-id="4d9d9-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-733">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-734">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-735">
         - File</span></span><br><span data-ttu-id="4d9d9-736">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-736">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-737">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-737">
         - Selection</span></span><br><span data-ttu-id="4d9d9-738">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-738">
         - Settings</span></span><br><span data-ttu-id="4d9d9-739">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-739">
         - TextCoercion</span></span><br><span data-ttu-id="4d9d9-740">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-740">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-741">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-741">Office 365 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-742">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-742">- Content</span></span><br><span data-ttu-id="4d9d9-743">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-743">
         - TaskPane</span></span><br><span data-ttu-id="4d9d9-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-744">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-745">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-746">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-746">- ActiveView</span></span><br><span data-ttu-id="4d9d9-747">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-747">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-748">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-748">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-749">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-749">
         - File</span></span><br><span data-ttu-id="4d9d9-750">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-750">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-751">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-751">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-752">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-752">
         - Selection</span></span><br><span data-ttu-id="4d9d9-753">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-753">
         - Settings</span></span><br><span data-ttu-id="4d9d9-754">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-754">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-755">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-755">Office 2019 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-756">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-756">- Content</span></span><br><span data-ttu-id="4d9d9-757">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-757">
         - TaskPane</span></span><br><span data-ttu-id="4d9d9-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-758">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-759">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-760">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-760">- ActiveView</span></span><br><span data-ttu-id="4d9d9-761">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-761">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-762">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-762">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-763">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-763">
         - File</span></span><br><span data-ttu-id="4d9d9-764">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-764">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-765">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-765">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-766">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-766">
         - Selection</span></span><br><span data-ttu-id="4d9d9-767">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-767">
         - Settings</span></span><br><span data-ttu-id="4d9d9-768">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-768">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-769">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="4d9d9-769">Office 2016 for Mac</span></span></td>
    <td> <span data-ttu-id="4d9d9-770">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-770">- Content</span></span><br><span data-ttu-id="4d9d9-771">
         - 作業ウィンドウ/td></span><span class="sxs-lookup"><span data-stu-id="4d9d9-771">
         - TaskPane/td></span></span> <td> <span data-ttu-id="4d9d9-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-772">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="4d9d9-773">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d9d9-773">- ActiveView</span></span><br><span data-ttu-id="4d9d9-774">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-774">
         - CompressedFile</span></span><br><span data-ttu-id="4d9d9-775">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-775">
         - DocumentEvents</span></span><br><span data-ttu-id="4d9d9-776">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d9d9-776">
         - File</span></span><br><span data-ttu-id="4d9d9-777">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-777">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-778">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d9d9-778">
         - PdfFile</span></span><br><span data-ttu-id="4d9d9-779">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-779">
         - Selection</span></span><br><span data-ttu-id="4d9d9-780">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-780">
         - Settings</span></span><br><span data-ttu-id="4d9d9-781">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-781">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d9d9-782">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d9d9-782">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4d9d9-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="4d9d9-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d9d9-784">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-784">Platform</span></span></th>
    <th><span data-ttu-id="4d9d9-785">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d9d9-785">Extension points</span></span></th>
    <th><span data-ttu-id="4d9d9-786">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-786">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d9d9-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-787"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="4d9d9-788">Office Online</span></span></td>
    <td> <span data-ttu-id="4d9d9-789">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-789">- Content</span></span><br><span data-ttu-id="4d9d9-790">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-790">
         - TaskPane</span></span><br><span data-ttu-id="4d9d9-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-791">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-792">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4d9d9-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-793">
         - <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-794">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d9d9-794">- DocumentEvents</span></span><br><span data-ttu-id="4d9d9-795">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-795">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d9d9-796">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-796">
         - ImageCoercion</span></span><br><span data-ttu-id="4d9d9-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d9d9-797">
         - Settings</span></span><br><span data-ttu-id="4d9d9-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-798">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4d9d9-799">Project</span><span class="sxs-lookup"><span data-stu-id="4d9d9-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d9d9-800">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d9d9-800">Platform</span></span></th>
    <th><span data-ttu-id="4d9d9-801">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d9d9-801">Extension points</span></span></th>
    <th><span data-ttu-id="4d9d9-802">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-802">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d9d9-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-803"><a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-804">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-804">Office 2019 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-805">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-805">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-806">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-807">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-807">- Selection</span></span><br><span data-ttu-id="4d9d9-808">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-808">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-809">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-809">Office 2016 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-810">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-810">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-811">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-812">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-812">- Selection</span></span><br><span data-ttu-id="4d9d9-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d9d9-814">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="4d9d9-814">Office 2013 for Windows</span></span></td>
    <td> <span data-ttu-id="4d9d9-815">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d9d9-815">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d9d9-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d9d9-816">- <a href="https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d9d9-817">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d9d9-817">- Selection</span></span><br><span data-ttu-id="4d9d9-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d9d9-818">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4d9d9-819">関連項目</span><span class="sxs-lookup"><span data-stu-id="4d9d9-819">See also</span></span>

- [<span data-ttu-id="4d9d9-820">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="4d9d9-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4d9d9-821">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-821">Common API requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="4d9d9-822">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="4d9d9-822">Add-in Commands requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="4d9d9-823">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="4d9d9-823">JavaScript API for Office reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)
