---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 05/23/2019
localization_priority: Priority
ms.openlocfilehash: 6fb1f0db839910e91d7a5215f8e21f5b33ff2165
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432195"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="5690b-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="5690b-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="5690b-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="5690b-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5690b-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="5690b-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="5690b-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="5690b-108">Excel</span><span class="sxs-lookup"><span data-stu-id="5690b-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5690b-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5690b-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5690b-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5690b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="5690b-113">Office Online</span></span></td>
    <td> <span data-ttu-id="5690b-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-114">- TaskPane</span></span><br><span data-ttu-id="5690b-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-115">
        - Content</span></span><br><span data-ttu-id="5690b-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-116">
        - Custom Functions</span></span><br><span data-ttu-id="5690b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="5690b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5690b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5690b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5690b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5690b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5690b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5690b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5690b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5690b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5690b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5690b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5690b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5690b-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-128">
        - BindingEvents</span></span><br><span data-ttu-id="5690b-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-129">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-130">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-131">
        - File</span></span><br><span data-ttu-id="5690b-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-132">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-134">
        - Selection</span></span><br><span data-ttu-id="5690b-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-135">
        - Settings</span></span><br><span data-ttu-id="5690b-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-136">
        - TableBindings</span></span><br><span data-ttu-id="5690b-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-137">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-138">
        - TextBindings</span></span><br><span data-ttu-id="5690b-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="5690b-140">Office on Windows</span></span><br><span data-ttu-id="5690b-141">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-142">- TaskPane</span></span><br><span data-ttu-id="5690b-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-143">
        - Content</span></span><br><span data-ttu-id="5690b-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-144">
        - Custom Functions</span></span><br><span data-ttu-id="5690b-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="5690b-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="5690b-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5690b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5690b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5690b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5690b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5690b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5690b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5690b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5690b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5690b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5690b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5690b-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-156">
        - BindingEvents</span></span><br><span data-ttu-id="5690b-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-157">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-158">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-159">
        - File</span></span><br><span data-ttu-id="5690b-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-160">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-162">
        - Selection</span></span><br><span data-ttu-id="5690b-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-163">
        - Settings</span></span><br><span data-ttu-id="5690b-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-164">
        - TableBindings</span></span><br><span data-ttu-id="5690b-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-165">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-166">
        - TextBindings</span></span><br><span data-ttu-id="5690b-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-168">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5690b-168">Office 2019 on Windows</span></span><br><span data-ttu-id="5690b-169">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5690b-170">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-170">- TaskPane</span></span><br><span data-ttu-id="5690b-171">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-171">
        - Content</span></span><br><span data-ttu-id="5690b-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5690b-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5690b-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5690b-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5690b-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5690b-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5690b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5690b-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5690b-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5690b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5690b-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-182">- BindingEvents</span></span><br><span data-ttu-id="5690b-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-183">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-184">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-185">
        - File</span></span><br><span data-ttu-id="5690b-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-186">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-187">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-189">
        - Selection</span></span><br><span data-ttu-id="5690b-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-190">
        - Settings</span></span><br><span data-ttu-id="5690b-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-191">
        - TableBindings</span></span><br><span data-ttu-id="5690b-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-192">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-193">
        - TextBindings</span></span><br><span data-ttu-id="5690b-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-195">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5690b-195">Office 2016 on Windows</span></span><br><span data-ttu-id="5690b-196">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5690b-197">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-197">- TaskPane</span></span><br><span data-ttu-id="5690b-198">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-198">
        - Content</span></span></td>
    <td><span data-ttu-id="5690b-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="5690b-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-201">- BindingEvents</span></span><br><span data-ttu-id="5690b-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-202">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-203">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-204">
        - File</span></span><br><span data-ttu-id="5690b-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-205">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-206">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-208">
        - Selection</span></span><br><span data-ttu-id="5690b-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-209">
        - Settings</span></span><br><span data-ttu-id="5690b-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-210">
        - TableBindings</span></span><br><span data-ttu-id="5690b-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-211">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-212">
        - TextBindings</span></span><br><span data-ttu-id="5690b-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-214">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="5690b-214">Office 2013 on Windows</span></span><br><span data-ttu-id="5690b-215">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5690b-216">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-216">
        - TaskPane</span></span><br><span data-ttu-id="5690b-217">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="5690b-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5690b-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="5690b-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-219">
        - BindingEvents</span></span><br><span data-ttu-id="5690b-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-220">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-221">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-222">
        - File</span></span><br><span data-ttu-id="5690b-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-223">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-224">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-226">
        - Selection</span></span><br><span data-ttu-id="5690b-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-227">
        - Settings</span></span><br><span data-ttu-id="5690b-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-228">
        - TableBindings</span></span><br><span data-ttu-id="5690b-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-229">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-230">
        - TextBindings</span></span><br><span data-ttu-id="5690b-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-232">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="5690b-232">Office for iPad</span></span><br><span data-ttu-id="5690b-233">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5690b-234">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-234">- TaskPane</span></span><br><span data-ttu-id="5690b-235">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-235">
        - Content</span></span><br><span data-ttu-id="5690b-236">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5690b-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5690b-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5690b-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5690b-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5690b-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5690b-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5690b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5690b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5690b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5690b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5690b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5690b-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-247">- BindingEvents</span></span><br><span data-ttu-id="5690b-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-248">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-249">
        - File</span></span><br><span data-ttu-id="5690b-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-250">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-251">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-253">
        - Selection</span></span><br><span data-ttu-id="5690b-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-254">
        - Settings</span></span><br><span data-ttu-id="5690b-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-255">
        - TableBindings</span></span><br><span data-ttu-id="5690b-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-256">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-257">
        - TextBindings</span></span><br><span data-ttu-id="5690b-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-259">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-259">Office for Mac</span></span><br><span data-ttu-id="5690b-260">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5690b-261">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-261">- TaskPane</span></span><br><span data-ttu-id="5690b-262">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-262">
        - Content</span></span><br><span data-ttu-id="5690b-263">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-263">
        - Custom Functions</span></span><br><span data-ttu-id="5690b-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5690b-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5690b-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5690b-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5690b-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5690b-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5690b-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5690b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5690b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5690b-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="5690b-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="5690b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5690b-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-275">- BindingEvents</span></span><br><span data-ttu-id="5690b-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-276">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-277">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-278">
        - File</span></span><br><span data-ttu-id="5690b-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-279">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-280">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-282">
        - PdfFile</span></span><br><span data-ttu-id="5690b-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-283">
        - Selection</span></span><br><span data-ttu-id="5690b-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-284">
        - Settings</span></span><br><span data-ttu-id="5690b-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-285">
        - TableBindings</span></span><br><span data-ttu-id="5690b-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-286">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-287">
        - TextBindings</span></span><br><span data-ttu-id="5690b-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-289">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-289">Office 2019 for Mac</span></span><br><span data-ttu-id="5690b-290">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5690b-291">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-291">- TaskPane</span></span><br><span data-ttu-id="5690b-292">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-292">
        - Content</span></span><br><span data-ttu-id="5690b-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="5690b-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="5690b-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="5690b-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="5690b-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="5690b-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="5690b-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="5690b-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="5690b-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="5690b-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="5690b-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-303">- BindingEvents</span></span><br><span data-ttu-id="5690b-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-304">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-305">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-306">
        - File</span></span><br><span data-ttu-id="5690b-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-307">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-308">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-310">
        - PdfFile</span></span><br><span data-ttu-id="5690b-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-311">
        - Selection</span></span><br><span data-ttu-id="5690b-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-312">
        - Settings</span></span><br><span data-ttu-id="5690b-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-313">
        - TableBindings</span></span><br><span data-ttu-id="5690b-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-314">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-315">
        - TextBindings</span></span><br><span data-ttu-id="5690b-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-317">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-317">Office 2016 for Mac</span></span><br><span data-ttu-id="5690b-318">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="5690b-319">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-319">- TaskPane</span></span><br><span data-ttu-id="5690b-320">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-320">
        - Content</span></span></td>
    <td><span data-ttu-id="5690b-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="5690b-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="5690b-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-323">- BindingEvents</span></span><br><span data-ttu-id="5690b-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-324">
        - CompressedFile</span></span><br><span data-ttu-id="5690b-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-325">
        - DocumentEvents</span></span><br><span data-ttu-id="5690b-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="5690b-326">
        - File</span></span><br><span data-ttu-id="5690b-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-327">
        - ImageCoercion</span></span><br><span data-ttu-id="5690b-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-328">
        - MatrixBindings</span></span><br><span data-ttu-id="5690b-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="5690b-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-330">
        - PdfFile</span></span><br><span data-ttu-id="5690b-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-331">
        - Selection</span></span><br><span data-ttu-id="5690b-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-332">
        - Settings</span></span><br><span data-ttu-id="5690b-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-333">
        - TableBindings</span></span><br><span data-ttu-id="5690b-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-334">
        - TableCoercion</span></span><br><span data-ttu-id="5690b-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-335">
        - TextBindings</span></span><br><span data-ttu-id="5690b-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5690b-337">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5690b-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="5690b-338">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="5690b-339">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="5690b-340">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="5690b-341">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="5690b-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="5690b-343">Office Online</span></span></td>
    <td><span data-ttu-id="5690b-344">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5690b-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-346">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="5690b-346">Office on Windows</span></span><br><span data-ttu-id="5690b-347">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5690b-348">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5690b-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-350">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="5690b-350">Office for iPad</span></span><br><span data-ttu-id="5690b-351">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5690b-352">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5690b-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-354">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-354">Office for Mac</span></span><br><span data-ttu-id="5690b-355">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="5690b-356">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="5690b-356">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="5690b-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="5690b-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="5690b-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5690b-359">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-359">Platform</span></span></th>
    <th><span data-ttu-id="5690b-360">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-360">Extension points</span></span></th>
    <th><span data-ttu-id="5690b-361">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="5690b-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="5690b-363">Office Online</span></span></td>
    <td> <span data-ttu-id="5690b-364">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-364">- Mail Read</span></span><br><span data-ttu-id="5690b-365">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-365">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5690b-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5690b-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5690b-374">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-375">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="5690b-375">Office on Windows</span></span><br><span data-ttu-id="5690b-376">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-377">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-377">- Mail Read</span></span><br><span data-ttu-id="5690b-378">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-378">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5690b-380">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5690b-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5690b-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5690b-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5690b-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5690b-388">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-389">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5690b-389">Office 2019 on Windows</span></span><br><span data-ttu-id="5690b-390">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-391">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-391">- Mail Read</span></span><br><span data-ttu-id="5690b-392">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-392">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5690b-394">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5690b-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5690b-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5690b-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5690b-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5690b-402">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-403">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5690b-403">Office 2016 on Windows</span></span><br><span data-ttu-id="5690b-404">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-405">- Mail Read</span></span><br><span data-ttu-id="5690b-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-406">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="5690b-408">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="5690b-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="5690b-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="5690b-413">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-414">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="5690b-414">Office 2013 on Windows</span></span><br><span data-ttu-id="5690b-415">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-416">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-416">- Mail Read</span></span><br><span data-ttu-id="5690b-417">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="5690b-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="5690b-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="5690b-422">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-423">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="5690b-423">Office for iOS</span></span><br><span data-ttu-id="5690b-424">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-425">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-425">- Mail Read</span></span><br><span data-ttu-id="5690b-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5690b-432">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-433">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-433">Office for Mac</span></span><br><span data-ttu-id="5690b-434">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-435">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-435">- Mail Read</span></span><br><span data-ttu-id="5690b-436">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-436">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5690b-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="5690b-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="5690b-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="5690b-445">使用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-446">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-446">Office 2019 for Mac</span></span><br><span data-ttu-id="5690b-447">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-448">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-448">- Mail Read</span></span><br><span data-ttu-id="5690b-449">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-449">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5690b-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5690b-457">利用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-458">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-458">Office 2016 for Mac</span></span><br><span data-ttu-id="5690b-459">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-460">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-460">- Mail Read</span></span><br><span data-ttu-id="5690b-461">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="5690b-461">
      - Mail Compose</span></span><br><span data-ttu-id="5690b-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="5690b-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="5690b-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="5690b-469">利用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-470">Office for Android</span><span class="sxs-lookup"><span data-stu-id="5690b-470">Office for Android</span></span><br><span data-ttu-id="5690b-471">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-471">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-472">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="5690b-472">- Mail Read</span></span><br><span data-ttu-id="5690b-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="5690b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="5690b-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="5690b-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="5690b-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="5690b-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="5690b-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="5690b-479">利用不可</span><span class="sxs-lookup"><span data-stu-id="5690b-479">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="5690b-480">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5690b-480">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="5690b-481">Word</span><span class="sxs-lookup"><span data-stu-id="5690b-481">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5690b-482">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-482">Platform</span></span></th>
    <th><span data-ttu-id="5690b-483">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-483">Extension points</span></span></th>
    <th><span data-ttu-id="5690b-484">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-484">API requirement sets</span></span></th>
    <th><span data-ttu-id="5690b-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-485"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-486">Office Online</span><span class="sxs-lookup"><span data-stu-id="5690b-486">Office Online</span></span></td>
    <td> <span data-ttu-id="5690b-487">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-487">- TaskPane</span></span><br><span data-ttu-id="5690b-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-488">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-489">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5690b-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5690b-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-492">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-493">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-493">- BindingEvents</span></span><br><span data-ttu-id="5690b-494">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-494">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-495">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-495">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-496">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-496">
         - File</span></span><br><span data-ttu-id="5690b-497">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-497">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-498">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-498">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-499">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-499">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-500">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-500">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-501">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-501">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-502">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-502">
         - PdfFile</span></span><br><span data-ttu-id="5690b-503">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-503">
         - Selection</span></span><br><span data-ttu-id="5690b-504">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-504">
         - Settings</span></span><br><span data-ttu-id="5690b-505">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-505">
         - TableBindings</span></span><br><span data-ttu-id="5690b-506">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-506">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-507">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-507">
         - TextBindings</span></span><br><span data-ttu-id="5690b-508">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-508">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-509">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-509">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-510">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="5690b-510">Office on Windows</span></span><br><span data-ttu-id="5690b-511">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-511">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-512">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-512">- TaskPane</span></span><br><span data-ttu-id="5690b-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-514">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5690b-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5690b-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-517">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-518">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-518">- BindingEvents</span></span><br><span data-ttu-id="5690b-519">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-519">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-520">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-520">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-521">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-521">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-522">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-522">
         - File</span></span><br><span data-ttu-id="5690b-523">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-523">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-524">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-524">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-525">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-525">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-526">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-526">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-527">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-527">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-528">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-528">
         - PdfFile</span></span><br><span data-ttu-id="5690b-529">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-529">
         - Selection</span></span><br><span data-ttu-id="5690b-530">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-530">
         - Settings</span></span><br><span data-ttu-id="5690b-531">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-531">
         - TableBindings</span></span><br><span data-ttu-id="5690b-532">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-532">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-533">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-533">
         - TextBindings</span></span><br><span data-ttu-id="5690b-534">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-534">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-535">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-535">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-536">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5690b-536">Office 2019 on Windows</span></span><br><span data-ttu-id="5690b-537">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-537">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-538">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-538">- TaskPane</span></span><br><span data-ttu-id="5690b-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-539">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-540">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5690b-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5690b-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-543">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-544">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-544">- BindingEvents</span></span><br><span data-ttu-id="5690b-545">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-545">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-546">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-546">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-547">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-547">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-548">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-548">
         - File</span></span><br><span data-ttu-id="5690b-549">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-549">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-550">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-550">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-551">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-551">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-552">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-552">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-553">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-553">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-554">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-554">
         - PdfFile</span></span><br><span data-ttu-id="5690b-555">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-555">
         - Selection</span></span><br><span data-ttu-id="5690b-556">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-556">
         - Settings</span></span><br><span data-ttu-id="5690b-557">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-557">
         - TableBindings</span></span><br><span data-ttu-id="5690b-558">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-558">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-559">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-559">
         - TextBindings</span></span><br><span data-ttu-id="5690b-560">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-560">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-561">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-561">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-562">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5690b-562">Office 2016 on Windows</span></span><br><span data-ttu-id="5690b-563">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-563">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-564">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-564">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-565">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-566">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="5690b-567">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-567">- BindingEvents</span></span><br><span data-ttu-id="5690b-568">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-568">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-569">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-569">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-570">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-570">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-571">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-571">
         - File</span></span><br><span data-ttu-id="5690b-572">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-572">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-573">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-573">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-574">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-577">
         - PdfFile</span></span><br><span data-ttu-id="5690b-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-578">
         - Selection</span></span><br><span data-ttu-id="5690b-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-579">
         - Settings</span></span><br><span data-ttu-id="5690b-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-580">
         - TableBindings</span></span><br><span data-ttu-id="5690b-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-581">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-582">
         - TextBindings</span></span><br><span data-ttu-id="5690b-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-583">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-585">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="5690b-585">Office 2013 on Windows</span></span><br><span data-ttu-id="5690b-586">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-587">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5690b-588">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5690b-589">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-589">- BindingEvents</span></span><br><span data-ttu-id="5690b-590">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-590">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-591">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-591">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-592">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-592">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-593">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-593">
         - File</span></span><br><span data-ttu-id="5690b-594">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-594">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-595">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-595">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-596">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-599">
         - PdfFile</span></span><br><span data-ttu-id="5690b-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-600">
         - Selection</span></span><br><span data-ttu-id="5690b-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-601">
         - Settings</span></span><br><span data-ttu-id="5690b-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-602">
         - TableBindings</span></span><br><span data-ttu-id="5690b-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-603">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-604">
         - TextBindings</span></span><br><span data-ttu-id="5690b-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-605">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-606">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-607">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="5690b-607">Office for iPad</span></span><br><span data-ttu-id="5690b-608">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-608">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-609">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5690b-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5690b-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5690b-613">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5690b-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-614">- BindingEvents</span></span><br><span data-ttu-id="5690b-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-615">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-617">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-618">
         - File</span></span><br><span data-ttu-id="5690b-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-620">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-620">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-621">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-621">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-622">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-622">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-623">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-623">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-624">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-624">
         - PdfFile</span></span><br><span data-ttu-id="5690b-625">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-625">
         - Selection</span></span><br><span data-ttu-id="5690b-626">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-626">
         - Settings</span></span><br><span data-ttu-id="5690b-627">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-627">
         - TableBindings</span></span><br><span data-ttu-id="5690b-628">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-628">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-629">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-629">
         - TextBindings</span></span><br><span data-ttu-id="5690b-630">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-630">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-631">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-631">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-632">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-632">Office for Mac</span></span><br><span data-ttu-id="5690b-633">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-633">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-634">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-634">- TaskPane</span></span><br><span data-ttu-id="5690b-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-636">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5690b-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5690b-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5690b-639">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5690b-640">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-640">- BindingEvents</span></span><br><span data-ttu-id="5690b-641">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-641">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-642">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-642">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-643">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-643">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-644">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-644">
         - File</span></span><br><span data-ttu-id="5690b-645">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-645">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-646">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-646">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-647">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-647">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-648">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-648">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-649">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-649">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-650">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-650">
         - PdfFile</span></span><br><span data-ttu-id="5690b-651">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-651">
         - Selection</span></span><br><span data-ttu-id="5690b-652">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-652">
         - Settings</span></span><br><span data-ttu-id="5690b-653">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-653">
         - TableBindings</span></span><br><span data-ttu-id="5690b-654">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-654">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-655">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-655">
         - TextBindings</span></span><br><span data-ttu-id="5690b-656">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-656">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-657">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-657">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-658">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-658">Office 2019 for Mac</span></span><br><span data-ttu-id="5690b-659">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-659">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-660">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-660">- TaskPane</span></span><br><span data-ttu-id="5690b-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-661">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-662">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="5690b-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="5690b-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="5690b-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="5690b-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="5690b-665">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="5690b-666">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-666">- BindingEvents</span></span><br><span data-ttu-id="5690b-667">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-667">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-668">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-668">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-669">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-669">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-670">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-670">
         - File</span></span><br><span data-ttu-id="5690b-671">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-671">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-672">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-672">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-673">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-673">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-674">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-674">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-675">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-675">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-676">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-676">
         - PdfFile</span></span><br><span data-ttu-id="5690b-677">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-677">
         - Selection</span></span><br><span data-ttu-id="5690b-678">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-678">
         - Settings</span></span><br><span data-ttu-id="5690b-679">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-679">
         - TableBindings</span></span><br><span data-ttu-id="5690b-680">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-680">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-681">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-681">
         - TextBindings</span></span><br><span data-ttu-id="5690b-682">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-682">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-683">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-683">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-684">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-684">Office 2016 for Mac</span></span><br><span data-ttu-id="5690b-685">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-685">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-686">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-686">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-687">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="5690b-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="5690b-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="5690b-689">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-689">- BindingEvents</span></span><br><span data-ttu-id="5690b-690">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-690">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-691">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="5690b-691">
         - CustomXmlParts</span></span><br><span data-ttu-id="5690b-692">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-692">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-693">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-693">
         - File</span></span><br><span data-ttu-id="5690b-694">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-694">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-695">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-695">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-696">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-696">
         - MatrixBindings</span></span><br><span data-ttu-id="5690b-697">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-697">
         - MatrixCoercion</span></span><br><span data-ttu-id="5690b-698">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-698">
         - OoxmlCoercion</span></span><br><span data-ttu-id="5690b-699">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-699">
         - PdfFile</span></span><br><span data-ttu-id="5690b-700">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-700">
         - Selection</span></span><br><span data-ttu-id="5690b-701">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-701">
         - Settings</span></span><br><span data-ttu-id="5690b-702">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-702">
         - TableBindings</span></span><br><span data-ttu-id="5690b-703">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-703">
         - TableCoercion</span></span><br><span data-ttu-id="5690b-704">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="5690b-704">
         - TextBindings</span></span><br><span data-ttu-id="5690b-705">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-705">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-706">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="5690b-706">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="5690b-707">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5690b-707">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="5690b-708">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="5690b-708">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5690b-709">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-709">Platform</span></span></th>
    <th><span data-ttu-id="5690b-710">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-710">Extension points</span></span></th>
    <th><span data-ttu-id="5690b-711">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-711">API requirement sets</span></span></th>
    <th><span data-ttu-id="5690b-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-712"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-713">Office Online</span><span class="sxs-lookup"><span data-stu-id="5690b-713">Office Online</span></span></td>
    <td> <span data-ttu-id="5690b-714">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-714">- Content</span></span><br><span data-ttu-id="5690b-715">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-715">
         - TaskPane</span></span><br><span data-ttu-id="5690b-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-716">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-717">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-718">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-718">- ActiveView</span></span><br><span data-ttu-id="5690b-719">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-719">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-720">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-720">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-721">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-721">
         - File</span></span><br><span data-ttu-id="5690b-722">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-722">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-723">
         - PdfFile</span></span><br><span data-ttu-id="5690b-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-724">
         - Selection</span></span><br><span data-ttu-id="5690b-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-725">
         - Settings</span></span><br><span data-ttu-id="5690b-726">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-726">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-727">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="5690b-727">Office on Windows</span></span><br><span data-ttu-id="5690b-728">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-728">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-729">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-729">- Content</span></span><br><span data-ttu-id="5690b-730">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-730">
         - TaskPane</span></span><br><span data-ttu-id="5690b-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-732">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-733">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-733">- ActiveView</span></span><br><span data-ttu-id="5690b-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-734">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-735">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-736">
         - File</span></span><br><span data-ttu-id="5690b-737">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-737">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-738">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-738">
         - PdfFile</span></span><br><span data-ttu-id="5690b-739">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-739">
         - Selection</span></span><br><span data-ttu-id="5690b-740">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-740">
         - Settings</span></span><br><span data-ttu-id="5690b-741">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-741">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-742">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5690b-742">Office 2019 on Windows</span></span><br><span data-ttu-id="5690b-743">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-743">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-744">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-744">- Content</span></span><br><span data-ttu-id="5690b-745">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-745">
         - TaskPane</span></span><br><span data-ttu-id="5690b-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-746">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-747">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-748">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-748">- ActiveView</span></span><br><span data-ttu-id="5690b-749">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-749">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-750">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-750">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-751">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-751">
         - File</span></span><br><span data-ttu-id="5690b-752">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-752">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-753">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-753">
         - PdfFile</span></span><br><span data-ttu-id="5690b-754">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-754">
         - Selection</span></span><br><span data-ttu-id="5690b-755">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-755">
         - Settings</span></span><br><span data-ttu-id="5690b-756">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-756">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-757">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5690b-757">Office 2016 on Windows</span></span><br><span data-ttu-id="5690b-758">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-758">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-759">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-759">- Content</span></span><br><span data-ttu-id="5690b-760">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-760">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5690b-761">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5690b-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-762">- ActiveView</span></span><br><span data-ttu-id="5690b-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-763">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-764">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-765">
         - File</span></span><br><span data-ttu-id="5690b-766">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-766">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-767">
         - PdfFile</span></span><br><span data-ttu-id="5690b-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-768">
         - Selection</span></span><br><span data-ttu-id="5690b-769">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-769">
         - Settings</span></span><br><span data-ttu-id="5690b-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-771">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="5690b-771">Office 2013 on Windows</span></span><br><span data-ttu-id="5690b-772">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-772">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-773">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-773">- Content</span></span><br><span data-ttu-id="5690b-774">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-774">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="5690b-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5690b-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5690b-776">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-776">- ActiveView</span></span><br><span data-ttu-id="5690b-777">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-777">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-778">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-778">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-779">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-779">
         - File</span></span><br><span data-ttu-id="5690b-780">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-780">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-781">
         - PdfFile</span></span><br><span data-ttu-id="5690b-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-782">
         - Selection</span></span><br><span data-ttu-id="5690b-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-783">
         - Settings</span></span><br><span data-ttu-id="5690b-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-785">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="5690b-785">Office for iPad</span></span><br><span data-ttu-id="5690b-786">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-786">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-787">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-787">- Content</span></span><br><span data-ttu-id="5690b-788">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-790">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-790">- ActiveView</span></span><br><span data-ttu-id="5690b-791">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-791">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-792">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-792">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-793">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-793">
         - File</span></span><br><span data-ttu-id="5690b-794">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-794">
         - PdfFile</span></span><br><span data-ttu-id="5690b-795">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-795">
         - Selection</span></span><br><span data-ttu-id="5690b-796">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-796">
         - Settings</span></span><br><span data-ttu-id="5690b-797">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-797">
         - TextCoercion</span></span><br><span data-ttu-id="5690b-798">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-798">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-799">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-799">Office for Mac</span></span><br><span data-ttu-id="5690b-800">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="5690b-800">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="5690b-801">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-801">- Content</span></span><br><span data-ttu-id="5690b-802">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-802">
         - TaskPane</span></span><br><span data-ttu-id="5690b-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-803">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-804">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-805">- ActiveView</span></span><br><span data-ttu-id="5690b-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-806">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-807">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-808">
         - File</span></span><br><span data-ttu-id="5690b-809">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-809">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-810">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-810">
         - PdfFile</span></span><br><span data-ttu-id="5690b-811">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-811">
         - Selection</span></span><br><span data-ttu-id="5690b-812">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-812">
         - Settings</span></span><br><span data-ttu-id="5690b-813">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-813">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-814">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-814">Office 2019 for Mac</span></span><br><span data-ttu-id="5690b-815">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-815">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-816">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-816">- Content</span></span><br><span data-ttu-id="5690b-817">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-817">
         - TaskPane</span></span><br><span data-ttu-id="5690b-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-819">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-820">- ActiveView</span></span><br><span data-ttu-id="5690b-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-821">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-822">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-823">
         - File</span></span><br><span data-ttu-id="5690b-824">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-824">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-825">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-825">
         - PdfFile</span></span><br><span data-ttu-id="5690b-826">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-826">
         - Selection</span></span><br><span data-ttu-id="5690b-827">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-827">
         - Settings</span></span><br><span data-ttu-id="5690b-828">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-828">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-829">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="5690b-829">Office 2016 for Mac</span></span><br><span data-ttu-id="5690b-830">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-830">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-831">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-831">- Content</span></span><br><span data-ttu-id="5690b-832">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-832">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="5690b-833">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="5690b-834">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="5690b-834">- ActiveView</span></span><br><span data-ttu-id="5690b-835">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="5690b-835">
         - CompressedFile</span></span><br><span data-ttu-id="5690b-836">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-836">
         - DocumentEvents</span></span><br><span data-ttu-id="5690b-837">
         - File</span><span class="sxs-lookup"><span data-stu-id="5690b-837">
         - File</span></span><br><span data-ttu-id="5690b-838">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-838">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-839">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="5690b-839">
         - PdfFile</span></span><br><span data-ttu-id="5690b-840">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-840">
         - Selection</span></span><br><span data-ttu-id="5690b-841">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-841">
         - Settings</span></span><br><span data-ttu-id="5690b-842">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-842">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="5690b-843">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="5690b-843">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="5690b-844">OneNote</span><span class="sxs-lookup"><span data-stu-id="5690b-844">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5690b-845">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-845">Platform</span></span></th>
    <th><span data-ttu-id="5690b-846">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-846">Extension points</span></span></th>
    <th><span data-ttu-id="5690b-847">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-847">API requirement sets</span></span></th>
    <th><span data-ttu-id="5690b-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-848"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-849">Office Online</span><span class="sxs-lookup"><span data-stu-id="5690b-849">Office Online</span></span></td>
    <td> <span data-ttu-id="5690b-850">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5690b-850">- Content</span></span><br><span data-ttu-id="5690b-851">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-851">
         - TaskPane</span></span><br><span data-ttu-id="5690b-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="5690b-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="5690b-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-853">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="5690b-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-855">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="5690b-855">- DocumentEvents</span></span><br><span data-ttu-id="5690b-856">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-856">
         - HtmlCoercion</span></span><br><span data-ttu-id="5690b-857">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-857">
         - ImageCoercion</span></span><br><span data-ttu-id="5690b-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="5690b-858">
         - Settings</span></span><br><span data-ttu-id="5690b-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-859">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="5690b-860">Project</span><span class="sxs-lookup"><span data-stu-id="5690b-860">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="5690b-861">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="5690b-861">Platform</span></span></th>
    <th><span data-ttu-id="5690b-862">拡張点</span><span class="sxs-lookup"><span data-stu-id="5690b-862">Extension points</span></span></th>
    <th><span data-ttu-id="5690b-863">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-863">API requirement sets</span></span></th>
    <th><span data-ttu-id="5690b-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="5690b-864"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-865">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="5690b-865">Office 2019 on Windows</span></span><br><span data-ttu-id="5690b-866">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-866">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-867">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-867">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-869">- Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-869">- Selection</span></span><br><span data-ttu-id="5690b-870">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-870">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-871">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="5690b-871">Office 2016 on Windows</span></span><br><span data-ttu-id="5690b-872">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-872">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-873">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-873">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-874">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-875">- Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-875">- Selection</span></span><br><span data-ttu-id="5690b-876">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-876">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="5690b-877">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="5690b-877">Office 2013 on Windows</span></span><br><span data-ttu-id="5690b-878">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="5690b-878">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="5690b-879">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5690b-879">- TaskPane</span></span></td>
    <td> <span data-ttu-id="5690b-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="5690b-880">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="5690b-881">- Selection</span><span class="sxs-lookup"><span data-stu-id="5690b-881">- Selection</span></span><br><span data-ttu-id="5690b-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="5690b-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="5690b-883">関連項目</span><span class="sxs-lookup"><span data-stu-id="5690b-883">See also</span></span>

- [<span data-ttu-id="5690b-884">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="5690b-884">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="5690b-885">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-885">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="5690b-886">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-886">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="5690b-887">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="5690b-887">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="5690b-888">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="5690b-888">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="5690b-889">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="5690b-889">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="5690b-890">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="5690b-890">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="5690b-891">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="5690b-891">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="5690b-892">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="5690b-892">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="5690b-893">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="5690b-893">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="5690b-894">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="5690b-894">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
