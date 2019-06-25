---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 06/13/2019
localization_priority: Priority
ms.openlocfilehash: 82c276c802cab66ae4f5443d0d556bc42ee57841
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128623"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1c270-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1c270-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="1c270-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="1c270-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="1c270-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="1c270-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="1c270-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="1c270-108">Excel</span><span class="sxs-lookup"><span data-stu-id="1c270-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1c270-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1c270-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1c270-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1c270-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="1c270-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-114">- TaskPane</span></span><br><span data-ttu-id="1c270-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-115">
        - Content</span></span><br><span data-ttu-id="1c270-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-116">
        - Custom Functions</span></span><br><span data-ttu-id="1c270-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="1c270-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1c270-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c270-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c270-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c270-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c270-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c270-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1c270-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1c270-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1c270-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1c270-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1c270-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c270-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-128">
        - BindingEvents</span></span><br><span data-ttu-id="1c270-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-129">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-130">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-131">
        - File</span></span><br><span data-ttu-id="1c270-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-132">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-134">
        - Selection</span></span><br><span data-ttu-id="1c270-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-135">
        - Settings</span></span><br><span data-ttu-id="1c270-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-136">
        - TableBindings</span></span><br><span data-ttu-id="1c270-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-137">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-138">
        - TextBindings</span></span><br><span data-ttu-id="1c270-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-140">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1c270-140">Office on Windows</span></span><br><span data-ttu-id="1c270-141">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-142">- TaskPane</span></span><br><span data-ttu-id="1c270-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-143">
        - Content</span></span><br><span data-ttu-id="1c270-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-144">
        - Custom Functions</span></span><br><span data-ttu-id="1c270-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="1c270-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1c270-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c270-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c270-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c270-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c270-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c270-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1c270-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1c270-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1c270-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1c270-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1c270-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c270-156">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-156">
        - BindingEvents</span></span><br><span data-ttu-id="1c270-157">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-157">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-158">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-158">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-159">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-159">
        - File</span></span><br><span data-ttu-id="1c270-160">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-160">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-161">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-161">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-162">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-162">
        - Selection</span></span><br><span data-ttu-id="1c270-163">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-163">
        - Settings</span></span><br><span data-ttu-id="1c270-164">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-164">
        - TableBindings</span></span><br><span data-ttu-id="1c270-165">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-165">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-166">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-166">
        - TextBindings</span></span><br><span data-ttu-id="1c270-167">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-167">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-168">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-168">Office 2019 on Windows</span></span><br><span data-ttu-id="1c270-169">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-169">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1c270-170">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-170">- TaskPane</span></span><br><span data-ttu-id="1c270-171">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-171">
        - Content</span></span><br><span data-ttu-id="1c270-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-172">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1c270-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-173">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c270-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-175">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c270-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c270-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c270-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c270-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1c270-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1c270-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1c270-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c270-182">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-182">- BindingEvents</span></span><br><span data-ttu-id="1c270-183">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-183">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-184">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-184">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-185">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-185">
        - File</span></span><br><span data-ttu-id="1c270-186">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-186">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-187">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-187">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-188">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-188">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-189">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-189">
        - Selection</span></span><br><span data-ttu-id="1c270-190">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-190">
        - Settings</span></span><br><span data-ttu-id="1c270-191">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-191">
        - TableBindings</span></span><br><span data-ttu-id="1c270-192">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-192">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-193">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-193">
        - TextBindings</span></span><br><span data-ttu-id="1c270-194">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-194">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-195">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-195">Office 2016 on Windows</span></span><br><span data-ttu-id="1c270-196">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-196">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1c270-197">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-197">- TaskPane</span></span><br><span data-ttu-id="1c270-198">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-198">
        - Content</span></span></td>
    <td><span data-ttu-id="1c270-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-199">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-200">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="1c270-201">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-201">- BindingEvents</span></span><br><span data-ttu-id="1c270-202">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-202">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-203">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-203">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-204">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-204">
        - File</span></span><br><span data-ttu-id="1c270-205">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-205">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-206">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-206">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-207">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-207">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-208">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-208">
        - Selection</span></span><br><span data-ttu-id="1c270-209">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-209">
        - Settings</span></span><br><span data-ttu-id="1c270-210">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-210">
        - TableBindings</span></span><br><span data-ttu-id="1c270-211">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-211">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-212">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-212">
        - TextBindings</span></span><br><span data-ttu-id="1c270-213">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-213">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-214">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1c270-214">Office 2013 on Windows</span></span><br><span data-ttu-id="1c270-215">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-215">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1c270-216">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-216">
        - TaskPane</span></span><br><span data-ttu-id="1c270-217">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-217">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1c270-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1c270-218">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td><span data-ttu-id="1c270-219">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-219">
        - BindingEvents</span></span><br><span data-ttu-id="1c270-220">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-220">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-221">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-221">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-222">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-222">
        - File</span></span><br><span data-ttu-id="1c270-223">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-223">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-224">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-224">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-225">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-225">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-226">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-226">
        - Selection</span></span><br><span data-ttu-id="1c270-227">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-227">
        - Settings</span></span><br><span data-ttu-id="1c270-228">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-228">
        - TableBindings</span></span><br><span data-ttu-id="1c270-229">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-229">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-230">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-230">
        - TextBindings</span></span><br><span data-ttu-id="1c270-231">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-231">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-232">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-232">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="1c270-233">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-233">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1c270-234">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-234">- TaskPane</span></span><br><span data-ttu-id="1c270-235">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-235">
        - Content</span></span><br><span data-ttu-id="1c270-236">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-236">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1c270-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-237">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-238">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c270-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c270-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c270-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c270-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c270-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1c270-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1c270-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1c270-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1c270-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1c270-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c270-247">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-247">- BindingEvents</span></span><br><span data-ttu-id="1c270-248">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-248">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-249">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-249">
        - File</span></span><br><span data-ttu-id="1c270-250">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-250">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-251">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-251">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-252">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-252">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-253">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-253">
        - Selection</span></span><br><span data-ttu-id="1c270-254">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-254">
        - Settings</span></span><br><span data-ttu-id="1c270-255">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-255">
        - TableBindings</span></span><br><span data-ttu-id="1c270-256">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-256">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-257">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-257">
        - TextBindings</span></span><br><span data-ttu-id="1c270-258">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-258">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-259">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-259">Office apps on Mac</span></span><br><span data-ttu-id="1c270-260">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-260">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1c270-261">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-261">- TaskPane</span></span><br><span data-ttu-id="1c270-262">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-262">
        - Content</span></span><br><span data-ttu-id="1c270-263">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-263">
        - Custom Functions</span></span><br><span data-ttu-id="1c270-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-264">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1c270-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-265">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c270-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c270-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c270-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c270-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c270-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1c270-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1c270-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1c270-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1c270-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1c270-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c270-275">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-275">- BindingEvents</span></span><br><span data-ttu-id="1c270-276">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-276">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-277">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-277">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-278">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-278">
        - File</span></span><br><span data-ttu-id="1c270-279">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-279">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-280">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-280">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-281">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-281">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-282">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-282">
        - PdfFile</span></span><br><span data-ttu-id="1c270-283">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-283">
        - Selection</span></span><br><span data-ttu-id="1c270-284">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-284">
        - Settings</span></span><br><span data-ttu-id="1c270-285">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-285">
        - TableBindings</span></span><br><span data-ttu-id="1c270-286">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-286">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-287">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-287">
        - TextBindings</span></span><br><span data-ttu-id="1c270-288">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-288">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-289">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-289">Office 2019 for Mac</span></span><br><span data-ttu-id="1c270-290">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-290">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1c270-291">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-291">- TaskPane</span></span><br><span data-ttu-id="1c270-292">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-292">
        - Content</span></span><br><span data-ttu-id="1c270-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-293">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1c270-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-294">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1c270-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1c270-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1c270-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1c270-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1c270-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1c270-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1c270-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1c270-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1c270-303">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-303">- BindingEvents</span></span><br><span data-ttu-id="1c270-304">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-304">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-305">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-305">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-306">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-306">
        - File</span></span><br><span data-ttu-id="1c270-307">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-307">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-308">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-308">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-309">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-309">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-310">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-310">
        - PdfFile</span></span><br><span data-ttu-id="1c270-311">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-311">
        - Selection</span></span><br><span data-ttu-id="1c270-312">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-312">
        - Settings</span></span><br><span data-ttu-id="1c270-313">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-313">
        - TableBindings</span></span><br><span data-ttu-id="1c270-314">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-314">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-315">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-315">
        - TextBindings</span></span><br><span data-ttu-id="1c270-316">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-316">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-317">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-317">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="1c270-318">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-318">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1c270-319">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-319">- TaskPane</span></span><br><span data-ttu-id="1c270-320">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-320">
        - Content</span></span></td>
    <td><span data-ttu-id="1c270-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-321">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1c270-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-322">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td><span data-ttu-id="1c270-323">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-323">- BindingEvents</span></span><br><span data-ttu-id="1c270-324">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-324">
        - CompressedFile</span></span><br><span data-ttu-id="1c270-325">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-325">
        - DocumentEvents</span></span><br><span data-ttu-id="1c270-326">
        - File</span><span class="sxs-lookup"><span data-stu-id="1c270-326">
        - File</span></span><br><span data-ttu-id="1c270-327">
        - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-327">
        - ImageCoercion</span></span><br><span data-ttu-id="1c270-328">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-328">
        - MatrixBindings</span></span><br><span data-ttu-id="1c270-329">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-329">
        - MatrixCoercion</span></span><br><span data-ttu-id="1c270-330">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-330">
        - PdfFile</span></span><br><span data-ttu-id="1c270-331">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-331">
        - Selection</span></span><br><span data-ttu-id="1c270-332">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-332">
        - Settings</span></span><br><span data-ttu-id="1c270-333">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-333">
        - TableBindings</span></span><br><span data-ttu-id="1c270-334">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-334">
        - TableCoercion</span></span><br><span data-ttu-id="1c270-335">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-335">
        - TextBindings</span></span><br><span data-ttu-id="1c270-336">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-336">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1c270-337">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1c270-337">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="1c270-338">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-338">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1c270-339">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-339">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1c270-340">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-340">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1c270-341">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-341">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1c270-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-342"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-343">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-343">Office on the web</span></span></td>
    <td><span data-ttu-id="1c270-344">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-344">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1c270-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-345">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-346">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1c270-346">Office on Windows</span></span><br><span data-ttu-id="1c270-347">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-347">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1c270-348">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-348">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1c270-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-349">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-350">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="1c270-350">Office for Mac</span></span><br><span data-ttu-id="1c270-351">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="1c270-351">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1c270-352">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1c270-352">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1c270-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-353">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="1c270-354">Outlook</span><span class="sxs-lookup"><span data-stu-id="1c270-354">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c270-355">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-355">Platform</span></span></th>
    <th><span data-ttu-id="1c270-356">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-356">Extension points</span></span></th>
    <th><span data-ttu-id="1c270-357">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-357">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c270-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-359">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-359">Office on the web</span></span><br><span data-ttu-id="1c270-360">(新規)</span><span class="sxs-lookup"><span data-stu-id="1c270-360">New</span></span></td>
    <td> <span data-ttu-id="1c270-361">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-361">- Mail Read</span></span><br><span data-ttu-id="1c270-362">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-362">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-363">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-364">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-365">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-366">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1c270-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1c270-371">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-371">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-372">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-372">Office on the web</span></span><br><span data-ttu-id="1c270-373">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="1c270-373">Classic</span></span></td>
    <td> <span data-ttu-id="1c270-374">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-374">- Mail Read</span></span><br><span data-ttu-id="1c270-375">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-375">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1c270-383">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-384">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1c270-384">Office on Windows</span></span><br><span data-ttu-id="1c270-385">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-385">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-386">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-386">- Mail Read</span></span><br><span data-ttu-id="1c270-387">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-387">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1c270-389">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="1c270-389">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1c270-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-390">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1c270-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1c270-397">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-397">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-398">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-398">Office 2019 on Windows</span></span><br><span data-ttu-id="1c270-399">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-399">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-400">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-400">- Mail Read</span></span><br><span data-ttu-id="1c270-401">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-401">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1c270-403">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="1c270-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1c270-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1c270-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1c270-411">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-411">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-412">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-412">Office 2016 on Windows</span></span><br><span data-ttu-id="1c270-413">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-413">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-414">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-414">- Mail Read</span></span><br><span data-ttu-id="1c270-415">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-415">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-416">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1c270-417">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="1c270-417">
      - Modules</span></span></td>
    <td> <span data-ttu-id="1c270-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-418">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1c270-422">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-422">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-423">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1c270-423">Office 2013 on Windows</span></span><br><span data-ttu-id="1c270-424">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-424">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-425">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-425">- Mail Read</span></span><br><span data-ttu-id="1c270-426">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-426">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="1c270-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="1c270-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1c270-431">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-432">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-432">Office apps on iOS</span></span><br><span data-ttu-id="1c270-433">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-433">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-434">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-434">- Mail Read</span></span><br><span data-ttu-id="1c270-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1c270-441">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-442">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-442">Office apps on Mac</span></span><br><span data-ttu-id="1c270-443">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-443">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-444">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-444">- Mail Read</span></span><br><span data-ttu-id="1c270-445">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-445">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-446">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-447">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1c270-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1c270-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1c270-454">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-454">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-455">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-455">Office 2019 for Mac</span></span><br><span data-ttu-id="1c270-456">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-456">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-457">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-457">- Mail Read</span></span><br><span data-ttu-id="1c270-458">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-458">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1c270-466">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-467">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-467">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="1c270-468">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-468">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-469">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-469">- Mail Read</span></span><br><span data-ttu-id="1c270-470">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="1c270-470">
      - Mail Compose</span></span><br><span data-ttu-id="1c270-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-471">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-472">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1c270-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1c270-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1c270-478">使用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-479">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-479">Office apps on Android</span></span><br><span data-ttu-id="1c270-480">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-480">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-481">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="1c270-481">- Mail Read</span></span><br><span data-ttu-id="1c270-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-482">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-483">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1c270-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1c270-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1c270-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1c270-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1c270-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1c270-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1c270-488">利用不可</span><span class="sxs-lookup"><span data-stu-id="1c270-488">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="1c270-489">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1c270-489">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="1c270-490">Word</span><span class="sxs-lookup"><span data-stu-id="1c270-490">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c270-491">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-491">Platform</span></span></th>
    <th><span data-ttu-id="1c270-492">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-492">Extension points</span></span></th>
    <th><span data-ttu-id="1c270-493">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-493">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c270-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-494"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-495">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-495">Office on the web</span></span></td>
    <td> <span data-ttu-id="1c270-496">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-496">- TaskPane</span></span><br><span data-ttu-id="1c270-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-497">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-498">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-499">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c270-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c270-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-502">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-502">- BindingEvents</span></span><br><span data-ttu-id="1c270-503">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-503">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-504">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-504">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-505">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-505">
         - File</span></span><br><span data-ttu-id="1c270-506">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-506">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-507">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-507">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-508">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-508">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-509">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-509">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-510">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-510">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-511">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-511">
         - PdfFile</span></span><br><span data-ttu-id="1c270-512">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-512">
         - Selection</span></span><br><span data-ttu-id="1c270-513">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-513">
         - Settings</span></span><br><span data-ttu-id="1c270-514">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-514">
         - TableBindings</span></span><br><span data-ttu-id="1c270-515">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-515">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-516">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-516">
         - TextBindings</span></span><br><span data-ttu-id="1c270-517">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-517">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-518">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-518">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-519">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1c270-519">Office on Windows</span></span><br><span data-ttu-id="1c270-520">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-520">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-521">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-521">- TaskPane</span></span><br><span data-ttu-id="1c270-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-522">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-523">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-524">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c270-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-525">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c270-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-526">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-527">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-527">- BindingEvents</span></span><br><span data-ttu-id="1c270-528">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-528">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-529">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-529">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-530">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-530">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-531">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-531">
         - File</span></span><br><span data-ttu-id="1c270-532">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-532">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-533">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-533">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-534">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-534">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-535">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-535">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-536">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-536">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-537">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-537">
         - PdfFile</span></span><br><span data-ttu-id="1c270-538">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-538">
         - Selection</span></span><br><span data-ttu-id="1c270-539">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-539">
         - Settings</span></span><br><span data-ttu-id="1c270-540">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-540">
         - TableBindings</span></span><br><span data-ttu-id="1c270-541">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-541">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-542">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-542">
         - TextBindings</span></span><br><span data-ttu-id="1c270-543">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-543">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-544">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-544">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-545">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-545">Office 2019 on Windows</span></span><br><span data-ttu-id="1c270-546">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-546">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-547">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-547">- TaskPane</span></span><br><span data-ttu-id="1c270-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-548">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-549">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-550">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c270-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-551">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c270-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-552">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-553">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-553">- BindingEvents</span></span><br><span data-ttu-id="1c270-554">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-554">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-555">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-555">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-556">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-556">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-557">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-557">
         - File</span></span><br><span data-ttu-id="1c270-558">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-558">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-559">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-559">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-560">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-560">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-561">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-561">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-562">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-562">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-563">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-563">
         - PdfFile</span></span><br><span data-ttu-id="1c270-564">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-564">
         - Selection</span></span><br><span data-ttu-id="1c270-565">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-565">
         - Settings</span></span><br><span data-ttu-id="1c270-566">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-566">
         - TableBindings</span></span><br><span data-ttu-id="1c270-567">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-567">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-568">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-568">
         - TextBindings</span></span><br><span data-ttu-id="1c270-569">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-569">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-570">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-570">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-571">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-571">Office 2016 on Windows</span></span><br><span data-ttu-id="1c270-572">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-572">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-573">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-573">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-574">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-575">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="1c270-576">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-576">- BindingEvents</span></span><br><span data-ttu-id="1c270-577">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-577">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-578">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-578">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-579">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-579">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-580">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-580">
         - File</span></span><br><span data-ttu-id="1c270-581">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-581">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-582">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-582">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-583">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-583">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-584">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-584">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-585">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-585">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-586">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-586">
         - PdfFile</span></span><br><span data-ttu-id="1c270-587">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-587">
         - Selection</span></span><br><span data-ttu-id="1c270-588">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-588">
         - Settings</span></span><br><span data-ttu-id="1c270-589">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-589">
         - TableBindings</span></span><br><span data-ttu-id="1c270-590">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-590">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-591">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-591">
         - TextBindings</span></span><br><span data-ttu-id="1c270-592">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-592">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-593">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-593">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-594">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1c270-594">Office 2013 on Windows</span></span><br><span data-ttu-id="1c270-595">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-595">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-596">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-596">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1c270-597">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1c270-598">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-598">- BindingEvents</span></span><br><span data-ttu-id="1c270-599">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-599">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-600">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-600">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-601">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-601">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-602">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-602">
         - File</span></span><br><span data-ttu-id="1c270-603">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-603">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-604">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-604">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-605">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-605">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-606">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-606">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-607">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-607">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-608">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-608">
         - PdfFile</span></span><br><span data-ttu-id="1c270-609">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-609">
         - Selection</span></span><br><span data-ttu-id="1c270-610">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-610">
         - Settings</span></span><br><span data-ttu-id="1c270-611">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-611">
         - TableBindings</span></span><br><span data-ttu-id="1c270-612">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-612">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-613">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-613">
         - TextBindings</span></span><br><span data-ttu-id="1c270-614">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-614">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-615">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-615">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-616">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-616">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="1c270-617">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-617">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-618">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-618">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-619">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-620">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c270-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-621">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c270-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1c270-622">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1c270-623">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-623">- BindingEvents</span></span><br><span data-ttu-id="1c270-624">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-624">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-625">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-625">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-626">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-626">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-627">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-627">
         - File</span></span><br><span data-ttu-id="1c270-628">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-628">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-629">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-629">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-630">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-630">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-631">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-631">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-632">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-632">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-633">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-633">
         - PdfFile</span></span><br><span data-ttu-id="1c270-634">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-634">
         - Selection</span></span><br><span data-ttu-id="1c270-635">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-635">
         - Settings</span></span><br><span data-ttu-id="1c270-636">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-636">
         - TableBindings</span></span><br><span data-ttu-id="1c270-637">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-637">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-638">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-638">
         - TextBindings</span></span><br><span data-ttu-id="1c270-639">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-639">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-640">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-640">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-641">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-641">Office apps on Mac</span></span><br><span data-ttu-id="1c270-642">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-642">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-643">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-643">- TaskPane</span></span><br><span data-ttu-id="1c270-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-644">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-645">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-646">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c270-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-647">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c270-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1c270-648">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1c270-649">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-649">- BindingEvents</span></span><br><span data-ttu-id="1c270-650">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-650">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-651">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-651">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-652">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-652">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-653">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-653">
         - File</span></span><br><span data-ttu-id="1c270-654">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-654">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-655">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-655">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-656">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-656">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-657">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-657">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-658">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-658">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-659">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-659">
         - PdfFile</span></span><br><span data-ttu-id="1c270-660">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-660">
         - Selection</span></span><br><span data-ttu-id="1c270-661">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-661">
         - Settings</span></span><br><span data-ttu-id="1c270-662">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-662">
         - TableBindings</span></span><br><span data-ttu-id="1c270-663">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-663">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-664">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-664">
         - TextBindings</span></span><br><span data-ttu-id="1c270-665">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-665">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-666">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-666">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-667">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-667">Office 2019 for Mac</span></span><br><span data-ttu-id="1c270-668">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-668">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-669">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-669">- TaskPane</span></span><br><span data-ttu-id="1c270-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-670">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-671">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1c270-672">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="1c270-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1c270-673">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="1c270-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span><span class="sxs-lookup"><span data-stu-id="1c270-674">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</span></span></td>
    <td> <span data-ttu-id="1c270-675">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-675">- BindingEvents</span></span><br><span data-ttu-id="1c270-676">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-676">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-677">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-677">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-678">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-678">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-679">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-679">
         - File</span></span><br><span data-ttu-id="1c270-680">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-680">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-681">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-681">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-682">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-682">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-683">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-683">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-684">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-684">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-685">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-685">
         - PdfFile</span></span><br><span data-ttu-id="1c270-686">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-686">
         - Selection</span></span><br><span data-ttu-id="1c270-687">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-687">
         - Settings</span></span><br><span data-ttu-id="1c270-688">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-688">
         - TableBindings</span></span><br><span data-ttu-id="1c270-689">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-689">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-690">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-690">
         - TextBindings</span></span><br><span data-ttu-id="1c270-691">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-691">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-692">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-692">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-693">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-693">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="1c270-694">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-694">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-695">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-695">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-696">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="1c270-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1c270-697">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span></td>
    <td> <span data-ttu-id="1c270-698">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-698">- BindingEvents</span></span><br><span data-ttu-id="1c270-699">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-699">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-700">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1c270-700">
         - CustomXmlParts</span></span><br><span data-ttu-id="1c270-701">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-701">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-702">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-702">
         - File</span></span><br><span data-ttu-id="1c270-703">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-703">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-704">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-704">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-705">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-705">
         - MatrixBindings</span></span><br><span data-ttu-id="1c270-706">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-706">
         - MatrixCoercion</span></span><br><span data-ttu-id="1c270-707">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-707">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1c270-708">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-708">
         - PdfFile</span></span><br><span data-ttu-id="1c270-709">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-709">
         - Selection</span></span><br><span data-ttu-id="1c270-710">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-710">
         - Settings</span></span><br><span data-ttu-id="1c270-711">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-711">
         - TableBindings</span></span><br><span data-ttu-id="1c270-712">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-712">
         - TableCoercion</span></span><br><span data-ttu-id="1c270-713">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1c270-713">
         - TextBindings</span></span><br><span data-ttu-id="1c270-714">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-714">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-715">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1c270-715">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="1c270-716">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1c270-716">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1c270-717">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1c270-717">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c270-718">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-718">Platform</span></span></th>
    <th><span data-ttu-id="1c270-719">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-719">Extension points</span></span></th>
    <th><span data-ttu-id="1c270-720">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-720">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c270-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-721"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-722">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-722">Office on the web</span></span></td>
    <td> <span data-ttu-id="1c270-723">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-723">- Content</span></span><br><span data-ttu-id="1c270-724">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-724">
         - TaskPane</span></span><br><span data-ttu-id="1c270-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-725">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-726">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-727">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-727">- ActiveView</span></span><br><span data-ttu-id="1c270-728">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-728">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-729">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-729">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-730">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-730">
         - File</span></span><br><span data-ttu-id="1c270-731">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-731">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-732">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-732">
         - PdfFile</span></span><br><span data-ttu-id="1c270-733">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-733">
         - Selection</span></span><br><span data-ttu-id="1c270-734">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-734">
         - Settings</span></span><br><span data-ttu-id="1c270-735">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-735">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-736">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1c270-736">Office on Windows</span></span><br><span data-ttu-id="1c270-737">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-737">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-738">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-738">- Content</span></span><br><span data-ttu-id="1c270-739">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-739">
         - TaskPane</span></span><br><span data-ttu-id="1c270-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-741">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-742">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-742">- ActiveView</span></span><br><span data-ttu-id="1c270-743">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-743">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-744">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-744">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-745">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-745">
         - File</span></span><br><span data-ttu-id="1c270-746">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-746">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-747">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-747">
         - PdfFile</span></span><br><span data-ttu-id="1c270-748">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-748">
         - Selection</span></span><br><span data-ttu-id="1c270-749">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-749">
         - Settings</span></span><br><span data-ttu-id="1c270-750">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-750">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-751">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-751">Office 2019 on Windows</span></span><br><span data-ttu-id="1c270-752">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-752">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-753">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-753">- Content</span></span><br><span data-ttu-id="1c270-754">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-754">
         - TaskPane</span></span><br><span data-ttu-id="1c270-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-756">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-757">- ActiveView</span></span><br><span data-ttu-id="1c270-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-758">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-759">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-760">
         - File</span></span><br><span data-ttu-id="1c270-761">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-761">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-762">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-762">
         - PdfFile</span></span><br><span data-ttu-id="1c270-763">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-763">
         - Selection</span></span><br><span data-ttu-id="1c270-764">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-764">
         - Settings</span></span><br><span data-ttu-id="1c270-765">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-765">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-766">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-766">Office 2016 on Windows</span></span><br><span data-ttu-id="1c270-767">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-767">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-768">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-768">- Content</span></span><br><span data-ttu-id="1c270-769">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-769">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1c270-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1c270-771">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-771">- ActiveView</span></span><br><span data-ttu-id="1c270-772">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-772">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-773">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-773">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-774">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-774">
         - File</span></span><br><span data-ttu-id="1c270-775">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-775">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-776">
         - PdfFile</span></span><br><span data-ttu-id="1c270-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-777">
         - Selection</span></span><br><span data-ttu-id="1c270-778">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-778">
         - Settings</span></span><br><span data-ttu-id="1c270-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-780">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1c270-780">Office 2013 on Windows</span></span><br><span data-ttu-id="1c270-781">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-782">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-782">- Content</span></span><br><span data-ttu-id="1c270-783">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-783">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="1c270-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1c270-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1c270-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-785">- ActiveView</span></span><br><span data-ttu-id="1c270-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-786">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-787">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-788">
         - File</span></span><br><span data-ttu-id="1c270-789">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-789">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-790">
         - PdfFile</span></span><br><span data-ttu-id="1c270-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-791">
         - Selection</span></span><br><span data-ttu-id="1c270-792">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-792">
         - Settings</span></span><br><span data-ttu-id="1c270-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-794">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-794">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="1c270-795">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-795">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-796">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-796">- Content</span></span><br><span data-ttu-id="1c270-797">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-797">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-799">- ActiveView</span></span><br><span data-ttu-id="1c270-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-800">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-801">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-802">
         - File</span></span><br><span data-ttu-id="1c270-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-803">
         - PdfFile</span></span><br><span data-ttu-id="1c270-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-804">
         - Selection</span></span><br><span data-ttu-id="1c270-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-805">
         - Settings</span></span><br><span data-ttu-id="1c270-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-806">
         - TextCoercion</span></span><br><span data-ttu-id="1c270-807">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-807">
         - ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-808">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1c270-808">Office apps on Mac</span></span><br><span data-ttu-id="1c270-809">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1c270-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1c270-810">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-810">- Content</span></span><br><span data-ttu-id="1c270-811">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-811">
         - TaskPane</span></span><br><span data-ttu-id="1c270-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-813">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-814">- ActiveView</span></span><br><span data-ttu-id="1c270-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-815">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-816">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-817">
         - File</span></span><br><span data-ttu-id="1c270-818">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-818">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-819">
         - PdfFile</span></span><br><span data-ttu-id="1c270-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-820">
         - Selection</span></span><br><span data-ttu-id="1c270-821">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-821">
         - Settings</span></span><br><span data-ttu-id="1c270-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-823">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-823">Office 2019 for Mac</span></span><br><span data-ttu-id="1c270-824">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-824">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-825">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-825">- Content</span></span><br><span data-ttu-id="1c270-826">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-826">
         - TaskPane</span></span><br><span data-ttu-id="1c270-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-828">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-829">- ActiveView</span></span><br><span data-ttu-id="1c270-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-830">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-831">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-832">
         - File</span></span><br><span data-ttu-id="1c270-833">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-833">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-834">
         - PdfFile</span></span><br><span data-ttu-id="1c270-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-835">
         - Selection</span></span><br><span data-ttu-id="1c270-836">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-836">
         - Settings</span></span><br><span data-ttu-id="1c270-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-838">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-838">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="1c270-839">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-840">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-840">- Content</span></span><br><span data-ttu-id="1c270-841">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-841">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1c270-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span></td>
    <td> <span data-ttu-id="1c270-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1c270-843">- ActiveView</span></span><br><span data-ttu-id="1c270-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1c270-844">
         - CompressedFile</span></span><br><span data-ttu-id="1c270-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-845">
         - DocumentEvents</span></span><br><span data-ttu-id="1c270-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="1c270-846">
         - File</span></span><br><span data-ttu-id="1c270-847">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-847">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1c270-848">
         - PdfFile</span></span><br><span data-ttu-id="1c270-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-849">
         - Selection</span></span><br><span data-ttu-id="1c270-850">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-850">
         - Settings</span></span><br><span data-ttu-id="1c270-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-851">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1c270-852">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1c270-852">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="1c270-853">OneNote</span><span class="sxs-lookup"><span data-stu-id="1c270-853">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c270-854">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-854">Platform</span></span></th>
    <th><span data-ttu-id="1c270-855">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-855">Extension points</span></span></th>
    <th><span data-ttu-id="1c270-856">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-856">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c270-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-857"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-858">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1c270-858">Office on the web</span></span></td>
    <td> <span data-ttu-id="1c270-859">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1c270-859">- Content</span></span><br><span data-ttu-id="1c270-860">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-860">
         - TaskPane</span></span><br><span data-ttu-id="1c270-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1c270-861">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1c270-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-862">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1c270-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-863">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-864">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1c270-864">- DocumentEvents</span></span><br><span data-ttu-id="1c270-865">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-865">
         - HtmlCoercion</span></span><br><span data-ttu-id="1c270-866">
         - ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-866">
         - ImageCoercion</span></span><br><span data-ttu-id="1c270-867">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1c270-867">
         - Settings</span></span><br><span data-ttu-id="1c270-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="1c270-869">Project</span><span class="sxs-lookup"><span data-stu-id="1c270-869">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1c270-870">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1c270-870">Platform</span></span></th>
    <th><span data-ttu-id="1c270-871">拡張点</span><span class="sxs-lookup"><span data-stu-id="1c270-871">Extension points</span></span></th>
    <th><span data-ttu-id="1c270-872">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-872">API requirement sets</span></span></th>
    <th><span data-ttu-id="1c270-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1c270-873"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-874">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1c270-874">Office 2019 on Windows</span></span><br><span data-ttu-id="1c270-875">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-875">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-876">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-876">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-877">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-878">- Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-878">- Selection</span></span><br><span data-ttu-id="1c270-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-879">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-880">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1c270-880">Office 2016 on Windows</span></span><br><span data-ttu-id="1c270-881">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-881">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-882">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-882">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-884">- Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-884">- Selection</span></span><br><span data-ttu-id="1c270-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-885">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1c270-886">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1c270-886">Office 2013 on Windows</span></span><br><span data-ttu-id="1c270-887">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1c270-887">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1c270-888">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1c270-888">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1c270-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1c270-889">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1c270-890">- Selection</span><span class="sxs-lookup"><span data-stu-id="1c270-890">- Selection</span></span><br><span data-ttu-id="1c270-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1c270-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1c270-892">関連項目</span><span class="sxs-lookup"><span data-stu-id="1c270-892">See also</span></span>

- [<span data-ttu-id="1c270-893">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="1c270-893">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1c270-894">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-894">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="1c270-895">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-895">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="1c270-896">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="1c270-896">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="1c270-897">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="1c270-897">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="1c270-898">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="1c270-898">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="1c270-899">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="1c270-899">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="1c270-900">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="1c270-900">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="1c270-901">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="1c270-901">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="1c270-902">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="1c270-902">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="1c270-903">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="1c270-903">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
