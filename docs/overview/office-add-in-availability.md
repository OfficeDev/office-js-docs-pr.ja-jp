---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 19f2fa7f744345823c2700b04524ec20705035a8
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952370"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="a12a5-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="a12a5-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a12a5-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="a12a5-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a12a5-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="a12a5-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="a12a5-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="a12a5-108">Excel</span><span class="sxs-lookup"><span data-stu-id="a12a5-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="a12a5-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="a12a5-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="a12a5-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="a12a5-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="a12a5-113">Office Online</span></span></td>
    <td> <span data-ttu-id="a12a5-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-114">- TaskPane</span></span><br><span data-ttu-id="a12a5-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-115">
        - Content</span></span><br><span data-ttu-id="a12a5-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-116">
        -Custom Functions</span></span><br><span data-ttu-id="a12a5-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="a12a5-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a12a5-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a12a5-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a12a5-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a12a5-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a12a5-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a12a5-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a12a5-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a12a5-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a12a5-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a12a5-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-128">
        - BindingEvents</span></span><br><span data-ttu-id="a12a5-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-129">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-130">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-131">
        - File</span></span><br><span data-ttu-id="a12a5-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-132">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-134">
        - Selection</span></span><br><span data-ttu-id="a12a5-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-135">
        - Settings</span></span><br><span data-ttu-id="a12a5-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-136">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-137">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-138">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-140">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="a12a5-140">Office apps on Windows</span></span><br><span data-ttu-id="a12a5-141">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-141">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-142">- TaskPane</span></span><br><span data-ttu-id="a12a5-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-143">
        - Content</span></span><br><span data-ttu-id="a12a5-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-144">
        -Custom Functions</span></span><br><span data-ttu-id="a12a5-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="a12a5-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a12a5-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a12a5-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a12a5-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a12a5-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a12a5-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a12a5-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a12a5-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a12a5-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a12a5-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a12a5-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-156">
        - BindingEvents</span></span><br><span data-ttu-id="a12a5-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-157">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-158">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-159">
        - File</span></span><br><span data-ttu-id="a12a5-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-160">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-162">
        - Selection</span></span><br><span data-ttu-id="a12a5-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-163">
        - Settings</span></span><br><span data-ttu-id="a12a5-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-164">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-165">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-166">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-168">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a12a5-168">Office 2019 for Windows</span></span><br><span data-ttu-id="a12a5-169">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a12a5-170">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-170">- TaskPane</span></span><br><span data-ttu-id="a12a5-171">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-171">
        - Content</span></span><br><span data-ttu-id="a12a5-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a12a5-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a12a5-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a12a5-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a12a5-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a12a5-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a12a5-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a12a5-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a12a5-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a12a5-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-182">- BindingEvents</span></span><br><span data-ttu-id="a12a5-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-183">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-184">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-185">
        - File</span></span><br><span data-ttu-id="a12a5-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-186">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-187">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-189">
        - Selection</span></span><br><span data-ttu-id="a12a5-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-190">
        - Settings</span></span><br><span data-ttu-id="a12a5-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-191">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-192">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-193">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-195">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a12a5-195">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="a12a5-196">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a12a5-197">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-197">- TaskPane</span></span><br><span data-ttu-id="a12a5-198">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-198">
        - Content</span></span></td>
    <td><span data-ttu-id="a12a5-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="a12a5-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-201">- BindingEvents</span></span><br><span data-ttu-id="a12a5-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-202">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-203">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-204">
        - File</span></span><br><span data-ttu-id="a12a5-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-205">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-206">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-208">
        - Selection</span></span><br><span data-ttu-id="a12a5-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-209">
        - Settings</span></span><br><span data-ttu-id="a12a5-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-210">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-211">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-212">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-214">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a12a5-214">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="a12a5-215">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a12a5-216">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-216">
        - TaskPane</span></span><br><span data-ttu-id="a12a5-217">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="a12a5-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a12a5-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="a12a5-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-219">
        - BindingEvents</span></span><br><span data-ttu-id="a12a5-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-220">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-221">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-222">
        - File</span></span><br><span data-ttu-id="a12a5-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-223">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-224">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-226">
        - Selection</span></span><br><span data-ttu-id="a12a5-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-227">
        - Settings</span></span><br><span data-ttu-id="a12a5-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-228">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-229">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-230">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-232">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="a12a5-232">Office for iPad</span></span><br><span data-ttu-id="a12a5-233">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-233">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="a12a5-234">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-234">- TaskPane</span></span><br><span data-ttu-id="a12a5-235">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-235">
        - Content</span></span><br><span data-ttu-id="a12a5-236">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-236">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="a12a5-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a12a5-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a12a5-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a12a5-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a12a5-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a12a5-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a12a5-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a12a5-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a12a5-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a12a5-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-247">- BindingEvents</span></span><br><span data-ttu-id="a12a5-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-248">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-249">
        - File</span></span><br><span data-ttu-id="a12a5-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-250">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-251">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-253">
        - Selection</span></span><br><span data-ttu-id="a12a5-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-254">
        - Settings</span></span><br><span data-ttu-id="a12a5-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-255">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-256">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-257">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-259">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-259">Office for Mac</span></span><br><span data-ttu-id="a12a5-260">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-260">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="a12a5-261">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-261">- TaskPane</span></span><br><span data-ttu-id="a12a5-262">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-262">
        - Content</span></span><br><span data-ttu-id="a12a5-263">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-263">
        -Custom Functions</span></span><br><span data-ttu-id="a12a5-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a12a5-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a12a5-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a12a5-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a12a5-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a12a5-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a12a5-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a12a5-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a12a5-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a12a5-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a12a5-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-275">- BindingEvents</span></span><br><span data-ttu-id="a12a5-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-276">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-277">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-278">
        - File</span></span><br><span data-ttu-id="a12a5-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-279">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-280">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-282">
        - PdfFile</span></span><br><span data-ttu-id="a12a5-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-283">
        - Selection</span></span><br><span data-ttu-id="a12a5-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-284">
        - Settings</span></span><br><span data-ttu-id="a12a5-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-285">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-286">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-287">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-289">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-289">Office 2019 for Mac</span></span><br><span data-ttu-id="a12a5-290">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a12a5-291">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-291">- TaskPane</span></span><br><span data-ttu-id="a12a5-292">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-292">
        - Content</span></span><br><span data-ttu-id="a12a5-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a12a5-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a12a5-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a12a5-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a12a5-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a12a5-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a12a5-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a12a5-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a12a5-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="a12a5-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-303">- BindingEvents</span></span><br><span data-ttu-id="a12a5-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-304">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-305">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-306">
        - File</span></span><br><span data-ttu-id="a12a5-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-307">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-308">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-310">
        - PdfFile</span></span><br><span data-ttu-id="a12a5-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-311">
        - Selection</span></span><br><span data-ttu-id="a12a5-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-312">
        - Settings</span></span><br><span data-ttu-id="a12a5-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-313">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-314">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-315">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-317">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-317">Office 2016 for Mac</span></span><br><span data-ttu-id="a12a5-318">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a12a5-319">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-319">- TaskPane</span></span><br><span data-ttu-id="a12a5-320">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-320">
        - Content</span></span></td>
    <td><span data-ttu-id="a12a5-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a12a5-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="a12a5-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-323">- BindingEvents</span></span><br><span data-ttu-id="a12a5-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-324">
        - CompressedFile</span></span><br><span data-ttu-id="a12a5-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-325">
        - DocumentEvents</span></span><br><span data-ttu-id="a12a5-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-326">
        - File</span></span><br><span data-ttu-id="a12a5-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-327">
        - ImageCoercion</span></span><br><span data-ttu-id="a12a5-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-328">
        - MatrixBindings</span></span><br><span data-ttu-id="a12a5-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-330">
        - PdfFile</span></span><br><span data-ttu-id="a12a5-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-331">
        - Selection</span></span><br><span data-ttu-id="a12a5-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-332">
        - Settings</span></span><br><span data-ttu-id="a12a5-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-333">
        - TableBindings</span></span><br><span data-ttu-id="a12a5-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-334">
        - TableCoercion</span></span><br><span data-ttu-id="a12a5-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-335">
        - TextBindings</span></span><br><span data-ttu-id="a12a5-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="a12a5-337">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a12a5-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="a12a5-338">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="a12a5-339">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="a12a5-340">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="a12a5-341">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="a12a5-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-343">Office Online</span><span class="sxs-lookup"><span data-stu-id="a12a5-343">Office Online</span></span></td>
    <td><span data-ttu-id="a12a5-344">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-344">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="a12a5-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-346">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="a12a5-346">Office apps on Windows</span></span><br><span data-ttu-id="a12a5-347">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-347">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="a12a5-348">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-348">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="a12a5-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-350">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="a12a5-350">Office for iPad</span></span><br><span data-ttu-id="a12a5-351">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="a12a5-352">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-352">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="a12a5-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-354">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-354">Office for Mac</span></span><br><span data-ttu-id="a12a5-355">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-355">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="a12a5-356">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a12a5-356">
        -Custom Functions</span></span></td>
    <td><span data-ttu-id="a12a5-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-357">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="a12a5-358">Outlook</span><span class="sxs-lookup"><span data-stu-id="a12a5-358">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a12a5-359">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-359">Platform</span></span></th>
    <th><span data-ttu-id="a12a5-360">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-360">Extension points</span></span></th>
    <th><span data-ttu-id="a12a5-361">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-361">API requirement sets</span></span></th>
    <th><span data-ttu-id="a12a5-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-362"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-363">Office Online</span><span class="sxs-lookup"><span data-stu-id="a12a5-363">Office Online</span></span></td>
    <td> <span data-ttu-id="a12a5-364">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-364">- Mail Read</span></span><br><span data-ttu-id="a12a5-365">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-365">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a12a5-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a12a5-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a12a5-374">使用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-375">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="a12a5-375">Office apps on Windows</span></span><br><span data-ttu-id="a12a5-376">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-376">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-377">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-377">- Mail Read</span></span><br><span data-ttu-id="a12a5-378">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-378">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a12a5-380">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="a12a5-380">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a12a5-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-381">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a12a5-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a12a5-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a12a5-388">使用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-389">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a12a5-389">Office 2019 for Windows</span></span><br><span data-ttu-id="a12a5-390">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-390">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-391">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-391">- Mail Read</span></span><br><span data-ttu-id="a12a5-392">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-392">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a12a5-394">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="a12a5-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a12a5-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a12a5-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a12a5-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a12a5-402">使用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-403">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a12a5-403">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="a12a5-404">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-405">- Mail Read</span></span><br><span data-ttu-id="a12a5-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-406">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a12a5-408">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="a12a5-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a12a5-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="a12a5-413">使用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-414">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a12a5-414">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="a12a5-415">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-416">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-416">- Mail Read</span></span><br><span data-ttu-id="a12a5-417">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-417">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="a12a5-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="a12a5-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="a12a5-422">使用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-423">Office for iOS</span><span class="sxs-lookup"><span data-stu-id="a12a5-423">Office for iOS</span></span><br><span data-ttu-id="a12a5-424">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-424">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-425">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-425">- Mail Read</span></span><br><span data-ttu-id="a12a5-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-426">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a12a5-432">使用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-433">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-433">Office for Mac</span></span><br><span data-ttu-id="a12a5-434">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-434">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-435">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-435">- Mail Read</span></span><br><span data-ttu-id="a12a5-436">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-436">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a12a5-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a12a5-444">利用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-445">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-445">Office 2019 for Mac</span></span><br><span data-ttu-id="a12a5-446">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-447">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-447">- Mail Read</span></span><br><span data-ttu-id="a12a5-448">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-448">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a12a5-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a12a5-456">利用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-457">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-457">Office 2016 for Mac</span></span><br><span data-ttu-id="a12a5-458">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-459">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-459">- Mail Read</span></span><br><span data-ttu-id="a12a5-460">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a12a5-460">
      - Mail Compose</span></span><br><span data-ttu-id="a12a5-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a12a5-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a12a5-468">利用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-469">Office for Android</span><span class="sxs-lookup"><span data-stu-id="a12a5-469">Office for Android</span></span><br><span data-ttu-id="a12a5-470">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-470">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-471">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a12a5-471">- Mail Read</span></span><br><span data-ttu-id="a12a5-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a12a5-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a12a5-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a12a5-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a12a5-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a12a5-478">利用不可</span><span class="sxs-lookup"><span data-stu-id="a12a5-478">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="a12a5-479">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a12a5-479">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="a12a5-480">Word</span><span class="sxs-lookup"><span data-stu-id="a12a5-480">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a12a5-481">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-481">Platform</span></span></th>
    <th><span data-ttu-id="a12a5-482">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-482">Extension points</span></span></th>
    <th><span data-ttu-id="a12a5-483">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-483">API requirement sets</span></span></th>
    <th><span data-ttu-id="a12a5-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-484"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-485">Office Online</span><span class="sxs-lookup"><span data-stu-id="a12a5-485">Office Online</span></span></td>
    <td> <span data-ttu-id="a12a5-486">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-486">- TaskPane</span></span><br><span data-ttu-id="a12a5-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-487">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-488">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-489">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a12a5-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-490">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a12a5-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-491">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-492">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-492">- BindingEvents</span></span><br><span data-ttu-id="a12a5-493">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-493">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-494">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-494">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-495">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-495">
         - File</span></span><br><span data-ttu-id="a12a5-496">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-496">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-497">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-497">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-498">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-498">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-499">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-499">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-500">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-500">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-501">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-501">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-502">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-502">
         - Selection</span></span><br><span data-ttu-id="a12a5-503">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-503">
         - Settings</span></span><br><span data-ttu-id="a12a5-504">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-504">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-505">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-505">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-506">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-506">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-507">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-507">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-508">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-508">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-509">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="a12a5-509">Office apps on Windows</span></span><br><span data-ttu-id="a12a5-510">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-510">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-511">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-511">- TaskPane</span></span><br><span data-ttu-id="a12a5-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-513">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-514">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a12a5-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-515">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a12a5-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-516">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-517">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-517">- BindingEvents</span></span><br><span data-ttu-id="a12a5-518">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-518">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-519">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-519">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-520">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-520">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-521">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-521">
         - File</span></span><br><span data-ttu-id="a12a5-522">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-522">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-523">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-523">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-524">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-524">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-525">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-525">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-526">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-526">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-527">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-527">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-528">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-528">
         - Selection</span></span><br><span data-ttu-id="a12a5-529">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-529">
         - Settings</span></span><br><span data-ttu-id="a12a5-530">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-530">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-531">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-531">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-532">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-532">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-533">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-533">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-534">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-534">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-535">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a12a5-535">Office 2019 for Windows</span></span><br><span data-ttu-id="a12a5-536">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-536">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-537">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-537">- TaskPane</span></span><br><span data-ttu-id="a12a5-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-539">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-540">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a12a5-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-541">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a12a5-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-542">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-543">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-543">- BindingEvents</span></span><br><span data-ttu-id="a12a5-544">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-544">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-545">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-545">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-546">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-546">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-547">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-547">
         - File</span></span><br><span data-ttu-id="a12a5-548">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-548">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-549">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-549">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-550">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-550">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-551">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-551">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-552">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-552">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-553">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-553">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-554">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-554">
         - Selection</span></span><br><span data-ttu-id="a12a5-555">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-555">
         - Settings</span></span><br><span data-ttu-id="a12a5-556">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-556">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-557">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-557">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-558">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-558">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-559">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-559">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-560">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-560">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-561">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a12a5-561">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="a12a5-562">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-562">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-563">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-563">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-564">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-565">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="a12a5-566">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-566">- BindingEvents</span></span><br><span data-ttu-id="a12a5-567">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-567">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-568">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-568">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-569">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-569">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-570">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-570">
         - File</span></span><br><span data-ttu-id="a12a5-571">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-571">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-572">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-572">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-573">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-573">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-574">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-574">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-575">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-575">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-576">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-576">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-577">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-577">
         - Selection</span></span><br><span data-ttu-id="a12a5-578">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-578">
         - Settings</span></span><br><span data-ttu-id="a12a5-579">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-579">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-580">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-580">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-581">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-581">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-582">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-582">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-583">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-583">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-584">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a12a5-584">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="a12a5-585">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-585">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-586">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-586">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a12a5-587">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a12a5-588">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-588">- BindingEvents</span></span><br><span data-ttu-id="a12a5-589">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-589">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-590">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-590">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-591">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-591">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-592">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-592">
         - File</span></span><br><span data-ttu-id="a12a5-593">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-593">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-594">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-594">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-595">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-595">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-596">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-596">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-597">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-597">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-598">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-598">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-599">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-599">
         - Selection</span></span><br><span data-ttu-id="a12a5-600">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-600">
         - Settings</span></span><br><span data-ttu-id="a12a5-601">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-601">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-602">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-602">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-603">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-603">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-604">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-604">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-605">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-605">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-606">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="a12a5-606">Office for iPad</span></span><br><span data-ttu-id="a12a5-607">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-607">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-608">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-608">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-609">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-610">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a12a5-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a12a5-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a12a5-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a12a5-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-613">- BindingEvents</span></span><br><span data-ttu-id="a12a5-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-614">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-616">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-617">
         - File</span></span><br><span data-ttu-id="a12a5-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-619">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-619">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-620">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-623">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-624">
         - Selection</span></span><br><span data-ttu-id="a12a5-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-625">
         - Settings</span></span><br><span data-ttu-id="a12a5-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-626">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-627">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-628">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-629">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-631">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-631">Office for Mac</span></span><br><span data-ttu-id="a12a5-632">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-632">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-633">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-633">- TaskPane</span></span><br><span data-ttu-id="a12a5-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-635">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-636">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a12a5-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-637">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a12a5-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a12a5-638">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a12a5-639">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-639">- BindingEvents</span></span><br><span data-ttu-id="a12a5-640">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-640">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-641">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-641">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-642">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-642">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-643">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-643">
         - File</span></span><br><span data-ttu-id="a12a5-644">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-644">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-645">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-645">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-646">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-646">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-647">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-647">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-648">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-648">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-649">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-649">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-650">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-650">
         - Selection</span></span><br><span data-ttu-id="a12a5-651">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-651">
         - Settings</span></span><br><span data-ttu-id="a12a5-652">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-652">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-653">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-653">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-654">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-654">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-655">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-655">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-656">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-656">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-657">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-657">Office 2019 for Mac</span></span><br><span data-ttu-id="a12a5-658">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-658">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-659">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-659">- TaskPane</span></span><br><span data-ttu-id="a12a5-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-661">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a12a5-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a12a5-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="a12a5-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="a12a5-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-665">- BindingEvents</span></span><br><span data-ttu-id="a12a5-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-666">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-668">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-669">
         - File</span></span><br><span data-ttu-id="a12a5-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-671">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-671">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-672">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-672">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-673">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-673">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-674">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-674">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-675">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-675">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-676">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-676">
         - Selection</span></span><br><span data-ttu-id="a12a5-677">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-677">
         - Settings</span></span><br><span data-ttu-id="a12a5-678">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-678">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-679">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-679">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-680">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-680">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-681">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-681">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-682">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-682">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-683">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-683">Office 2016 for Mac</span></span><br><span data-ttu-id="a12a5-684">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-684">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-685">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-685">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a12a5-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a12a5-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="a12a5-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-688">- BindingEvents</span></span><br><span data-ttu-id="a12a5-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-689">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a12a5-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="a12a5-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-691">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-692">
         - File</span></span><br><span data-ttu-id="a12a5-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-694">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-694">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-695">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-695">
         - MatrixBindings</span></span><br><span data-ttu-id="a12a5-696">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-696">
         - MatrixCoercion</span></span><br><span data-ttu-id="a12a5-697">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-697">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a12a5-698">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-698">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-699">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-699">
         - Selection</span></span><br><span data-ttu-id="a12a5-700">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-700">
         - Settings</span></span><br><span data-ttu-id="a12a5-701">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-701">
         - TableBindings</span></span><br><span data-ttu-id="a12a5-702">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-702">
         - TableCoercion</span></span><br><span data-ttu-id="a12a5-703">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a12a5-703">
         - TextBindings</span></span><br><span data-ttu-id="a12a5-704">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-704">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-705">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-705">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="a12a5-706">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a12a5-706">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="a12a5-707">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a12a5-707">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a12a5-708">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-708">Platform</span></span></th>
    <th><span data-ttu-id="a12a5-709">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-709">Extension points</span></span></th>
    <th><span data-ttu-id="a12a5-710">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-710">API requirement sets</span></span></th>
    <th><span data-ttu-id="a12a5-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-711"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-712">Office Online</span><span class="sxs-lookup"><span data-stu-id="a12a5-712">Office Online</span></span></td>
    <td> <span data-ttu-id="a12a5-713">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-713">- Content</span></span><br><span data-ttu-id="a12a5-714">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-714">
         - TaskPane</span></span><br><span data-ttu-id="a12a5-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-715">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-716">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-717">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-717">- ActiveView</span></span><br><span data-ttu-id="a12a5-718">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-718">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-719">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-719">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-720">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-720">
         - File</span></span><br><span data-ttu-id="a12a5-721">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-721">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-722">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-723">
         - Selection</span></span><br><span data-ttu-id="a12a5-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-724">
         - Settings</span></span><br><span data-ttu-id="a12a5-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-725">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-726">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="a12a5-726">Office apps on Windows</span></span><br><span data-ttu-id="a12a5-727">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-727">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-728">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-728">- Content</span></span><br><span data-ttu-id="a12a5-729">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-729">
         - TaskPane</span></span><br><span data-ttu-id="a12a5-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-730">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-731">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-732">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-732">- ActiveView</span></span><br><span data-ttu-id="a12a5-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-733">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-734">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-734">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-735">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-735">
         - File</span></span><br><span data-ttu-id="a12a5-736">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-736">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-737">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-737">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-738">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-738">
         - Selection</span></span><br><span data-ttu-id="a12a5-739">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-739">
         - Settings</span></span><br><span data-ttu-id="a12a5-740">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-740">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-741">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a12a5-741">Office 2019 for Windows</span></span><br><span data-ttu-id="a12a5-742">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-742">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-743">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-743">- Content</span></span><br><span data-ttu-id="a12a5-744">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-744">
         - TaskPane</span></span><br><span data-ttu-id="a12a5-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-745">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-746">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-747">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-747">- ActiveView</span></span><br><span data-ttu-id="a12a5-748">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-748">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-749">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-749">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-750">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-750">
         - File</span></span><br><span data-ttu-id="a12a5-751">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-751">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-752">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-752">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-753">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-753">
         - Selection</span></span><br><span data-ttu-id="a12a5-754">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-754">
         - Settings</span></span><br><span data-ttu-id="a12a5-755">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-755">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-756">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a12a5-756">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="a12a5-757">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-757">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-758">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-758">- Content</span></span><br><span data-ttu-id="a12a5-759">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-759">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a12a5-760">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a12a5-761">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-761">- ActiveView</span></span><br><span data-ttu-id="a12a5-762">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-762">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-763">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-763">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-764">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-764">
         - File</span></span><br><span data-ttu-id="a12a5-765">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-765">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-766">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-767">
         - Selection</span></span><br><span data-ttu-id="a12a5-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-768">
         - Settings</span></span><br><span data-ttu-id="a12a5-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-770">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a12a5-770">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="a12a5-771">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-772">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-772">- Content</span></span><br><span data-ttu-id="a12a5-773">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-773">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="a12a5-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a12a5-774">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a12a5-775">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-775">- ActiveView</span></span><br><span data-ttu-id="a12a5-776">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-776">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-777">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-777">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-778">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-778">
         - File</span></span><br><span data-ttu-id="a12a5-779">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-779">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-780">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-780">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-781">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-781">
         - Selection</span></span><br><span data-ttu-id="a12a5-782">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-782">
         - Settings</span></span><br><span data-ttu-id="a12a5-783">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-783">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-784">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="a12a5-784">Office for iPad</span></span><br><span data-ttu-id="a12a5-785">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-785">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-786">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-786">- Content</span></span><br><span data-ttu-id="a12a5-787">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-787">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-788">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-789">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-789">- ActiveView</span></span><br><span data-ttu-id="a12a5-790">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-790">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-791">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-791">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-792">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-792">
         - File</span></span><br><span data-ttu-id="a12a5-793">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-793">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-794">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-794">
         - Selection</span></span><br><span data-ttu-id="a12a5-795">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-795">
         - Settings</span></span><br><span data-ttu-id="a12a5-796">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-796">
         - TextCoercion</span></span><br><span data-ttu-id="a12a5-797">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-797">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-798">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-798">Office for Mac</span></span><br><span data-ttu-id="a12a5-799">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a12a5-799">(connected to Office 365)</span></span></td>
    <td> <span data-ttu-id="a12a5-800">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-800">- Content</span></span><br><span data-ttu-id="a12a5-801">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-801">
         - TaskPane</span></span><br><span data-ttu-id="a12a5-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-802">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-804">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-804">- ActiveView</span></span><br><span data-ttu-id="a12a5-805">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-805">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-806">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-806">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-807">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-807">
         - File</span></span><br><span data-ttu-id="a12a5-808">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-808">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-809">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-810">
         - Selection</span></span><br><span data-ttu-id="a12a5-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-811">
         - Settings</span></span><br><span data-ttu-id="a12a5-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-813">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-813">Office 2019 for Mac</span></span><br><span data-ttu-id="a12a5-814">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-814">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-815">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-815">- Content</span></span><br><span data-ttu-id="a12a5-816">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-816">
         - TaskPane</span></span><br><span data-ttu-id="a12a5-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-817">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-818">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-819">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-819">- ActiveView</span></span><br><span data-ttu-id="a12a5-820">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-820">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-821">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-821">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-822">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-822">
         - File</span></span><br><span data-ttu-id="a12a5-823">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-823">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-824">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-825">
         - Selection</span></span><br><span data-ttu-id="a12a5-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-826">
         - Settings</span></span><br><span data-ttu-id="a12a5-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-828">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="a12a5-828">Office 2016 for Mac</span></span><br><span data-ttu-id="a12a5-829">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-829">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-830">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-830">- Content</span></span><br><span data-ttu-id="a12a5-831">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-831">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a12a5-832">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="a12a5-833">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a12a5-833">- ActiveView</span></span><br><span data-ttu-id="a12a5-834">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-834">
         - CompressedFile</span></span><br><span data-ttu-id="a12a5-835">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-835">
         - DocumentEvents</span></span><br><span data-ttu-id="a12a5-836">
         - File</span><span class="sxs-lookup"><span data-stu-id="a12a5-836">
         - File</span></span><br><span data-ttu-id="a12a5-837">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-837">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-838">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a12a5-838">
         - PdfFile</span></span><br><span data-ttu-id="a12a5-839">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-839">
         - Selection</span></span><br><span data-ttu-id="a12a5-840">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-840">
         - Settings</span></span><br><span data-ttu-id="a12a5-841">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-841">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="a12a5-842">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a12a5-842">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="a12a5-843">OneNote</span><span class="sxs-lookup"><span data-stu-id="a12a5-843">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a12a5-844">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-844">Platform</span></span></th>
    <th><span data-ttu-id="a12a5-845">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-845">Extension points</span></span></th>
    <th><span data-ttu-id="a12a5-846">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-846">API requirement sets</span></span></th>
    <th><span data-ttu-id="a12a5-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-847"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-848">Office Online</span><span class="sxs-lookup"><span data-stu-id="a12a5-848">Office Online</span></span></td>
    <td> <span data-ttu-id="a12a5-849">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a12a5-849">- Content</span></span><br><span data-ttu-id="a12a5-850">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-850">
         - TaskPane</span></span><br><span data-ttu-id="a12a5-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a12a5-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-852">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="a12a5-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-854">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a12a5-854">- DocumentEvents</span></span><br><span data-ttu-id="a12a5-855">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-855">
         - HtmlCoercion</span></span><br><span data-ttu-id="a12a5-856">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-856">
         - ImageCoercion</span></span><br><span data-ttu-id="a12a5-857">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a12a5-857">
         - Settings</span></span><br><span data-ttu-id="a12a5-858">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-858">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="a12a5-859">Project</span><span class="sxs-lookup"><span data-stu-id="a12a5-859">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a12a5-860">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a12a5-860">Platform</span></span></th>
    <th><span data-ttu-id="a12a5-861">拡張点</span><span class="sxs-lookup"><span data-stu-id="a12a5-861">Extension points</span></span></th>
    <th><span data-ttu-id="a12a5-862">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-862">API requirement sets</span></span></th>
    <th><span data-ttu-id="a12a5-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a12a5-863"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-864">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a12a5-864">Office 2019 for Windows</span></span><br><span data-ttu-id="a12a5-865">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-866">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-866">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-867">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-868">- Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-868">- Selection</span></span><br><span data-ttu-id="a12a5-869">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-869">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-870">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a12a5-870">Set up Office 2016 on Windows Phone 8</span></span><br><span data-ttu-id="a12a5-871">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-871">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-872">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-872">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-874">- Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-874">- Selection</span></span><br><span data-ttu-id="a12a5-875">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-875">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a12a5-876">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a12a5-876">Enable Modern Authentication for Office 2013 on Windows devices</span></span><br><span data-ttu-id="a12a5-877">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a12a5-877">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a12a5-878">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a12a5-878">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a12a5-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a12a5-879">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a12a5-880">- Selection</span><span class="sxs-lookup"><span data-stu-id="a12a5-880">- Selection</span></span><br><span data-ttu-id="a12a5-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a12a5-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="a12a5-882">関連項目</span><span class="sxs-lookup"><span data-stu-id="a12a5-882">See also</span></span>

- [<span data-ttu-id="a12a5-883">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="a12a5-883">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="a12a5-884">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-884">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a12a5-885">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-885">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="a12a5-886">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="a12a5-886">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="a12a5-887">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="a12a5-887">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="a12a5-888">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="a12a5-888">Update history for Office 365 ProPlus releases</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="a12a5-889">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="a12a5-889">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="a12a5-890">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="a12a5-890">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="a12a5-891">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="a12a5-891">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="a12a5-892">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="a12a5-892">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="a12a5-893">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="a12a5-893">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
