---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: ecb906e595c08b973b5146416a5317d59547ed39
ms.sourcegitcommit: e56bd8f1260c73daf33272a30dc5af242452594f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/21/2019
ms.locfileid: "38757486"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="4d90c-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="4d90c-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d90c-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="4d90c-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4d90c-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="4d90c-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="4d90c-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="4d90c-108">Excel</span><span class="sxs-lookup"><span data-stu-id="4d90c-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4d90c-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4d90c-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4d90c-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4d90c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="4d90c-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-114">- TaskPane</span></span><br><span data-ttu-id="4d90c-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-115">
        - Content</span></span><br><span data-ttu-id="4d90c-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-116">
        - Custom Functions</span></span><br><span data-ttu-id="4d90c-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4d90c-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4d90c-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d90c-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d90c-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d90c-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d90c-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d90c-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d90c-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d90c-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4d90c-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4d90c-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="4d90c-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-130">
        - BindingEvents</span></span><br><span data-ttu-id="4d90c-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-131">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-132">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-133">
        - File</span></span><br><span data-ttu-id="4d90c-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-134">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-136">
        - Selection</span></span><br><span data-ttu-id="4d90c-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-137">
        - Settings</span></span><br><span data-ttu-id="4d90c-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-138">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-139">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-140">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-142">Office on Windows</span></span><br><span data-ttu-id="4d90c-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-144">- TaskPane</span></span><br><span data-ttu-id="4d90c-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-145">
        - Content</span></span><br><span data-ttu-id="4d90c-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-146">
        - Custom Functions</span></span><br><span data-ttu-id="4d90c-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4d90c-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4d90c-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d90c-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d90c-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d90c-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d90c-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d90c-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d90c-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d90c-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4d90c-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4d90c-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="4d90c-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-161">
        - BindingEvents</span></span><br><span data-ttu-id="4d90c-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-162">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-163">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-164">
        - File</span></span><br><span data-ttu-id="4d90c-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-165">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-167">
        - Selection</span></span><br><span data-ttu-id="4d90c-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-168">
        - Settings</span></span><br><span data-ttu-id="4d90c-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-169">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-170">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-171">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-173">Office 2019 on Windows</span></span><br><span data-ttu-id="4d90c-174">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4d90c-175">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-175">- TaskPane</span></span><br><span data-ttu-id="4d90c-176">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-176">
        - Content</span></span><br><span data-ttu-id="4d90c-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d90c-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d90c-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d90c-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d90c-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d90c-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d90c-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d90c-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d90c-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-188">- BindingEvents</span></span><br><span data-ttu-id="4d90c-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-189">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-190">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-191">
        - File</span></span><br><span data-ttu-id="4d90c-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-192">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-194">
        - Selection</span></span><br><span data-ttu-id="4d90c-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-195">
        - Settings</span></span><br><span data-ttu-id="4d90c-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-196">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-197">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-198">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-200">Office 2016 on Windows</span></span><br><span data-ttu-id="4d90c-201">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4d90c-202">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-202">- TaskPane</span></span><br><span data-ttu-id="4d90c-203">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-203">
        - Content</span></span></td>
    <td><span data-ttu-id="4d90c-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4d90c-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-207">- BindingEvents</span></span><br><span data-ttu-id="4d90c-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-208">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-209">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-210">
        - File</span></span><br><span data-ttu-id="4d90c-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-211">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-213">
        - Selection</span></span><br><span data-ttu-id="4d90c-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-214">
        - Settings</span></span><br><span data-ttu-id="4d90c-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-215">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-216">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-217">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4d90c-219">Office 2013 on Windows</span></span><br><span data-ttu-id="4d90c-220">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4d90c-221">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-221">
        - TaskPane</span></span><br><span data-ttu-id="4d90c-222">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="4d90c-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d90c-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4d90c-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-225">
        - BindingEvents</span></span><br><span data-ttu-id="4d90c-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-226">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-227">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-228">
        - File</span></span><br><span data-ttu-id="4d90c-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-229">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-231">
        - Selection</span></span><br><span data-ttu-id="4d90c-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-232">
        - Settings</span></span><br><span data-ttu-id="4d90c-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-233">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-234">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-235">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-237">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-237">Office on iPad</span></span><br><span data-ttu-id="4d90c-238">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="4d90c-239">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-239">- TaskPane</span></span><br><span data-ttu-id="4d90c-240">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-240">
        - Content</span></span></td>
    <td><span data-ttu-id="4d90c-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d90c-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d90c-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d90c-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d90c-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d90c-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d90c-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d90c-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4d90c-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4d90c-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-253">- BindingEvents</span></span><br><span data-ttu-id="4d90c-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-254">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-255">
        - File</span></span><br><span data-ttu-id="4d90c-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-256">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-258">
        - Selection</span></span><br><span data-ttu-id="4d90c-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-259">
        - Settings</span></span><br><span data-ttu-id="4d90c-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-260">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-261">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-262">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-264">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-264">Office on Mac</span></span><br><span data-ttu-id="4d90c-265">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="4d90c-266">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-266">- TaskPane</span></span><br><span data-ttu-id="4d90c-267">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-267">
        - Content</span></span><br><span data-ttu-id="4d90c-268">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-268">
        - Custom Functions</span></span><br><span data-ttu-id="4d90c-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d90c-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d90c-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d90c-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d90c-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d90c-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d90c-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d90c-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d90c-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4d90c-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4d90c-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="4d90c-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-283">- BindingEvents</span></span><br><span data-ttu-id="4d90c-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-284">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-285">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-286">
        - File</span></span><br><span data-ttu-id="4d90c-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-287">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-289">
        - PdfFile</span></span><br><span data-ttu-id="4d90c-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-290">
        - Selection</span></span><br><span data-ttu-id="4d90c-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-291">
        - Settings</span></span><br><span data-ttu-id="4d90c-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-292">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-293">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-294">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-296">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-296">Office 2019 on Mac</span></span><br><span data-ttu-id="4d90c-297">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4d90c-298">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-298">- TaskPane</span></span><br><span data-ttu-id="4d90c-299">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-299">
        - Content</span></span><br><span data-ttu-id="4d90c-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="4d90c-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4d90c-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4d90c-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4d90c-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4d90c-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4d90c-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4d90c-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4d90c-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-311">- BindingEvents</span></span><br><span data-ttu-id="4d90c-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-312">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-313">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-314">
        - File</span></span><br><span data-ttu-id="4d90c-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-315">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-317">
        - PdfFile</span></span><br><span data-ttu-id="4d90c-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-318">
        - Selection</span></span><br><span data-ttu-id="4d90c-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-319">
        - Settings</span></span><br><span data-ttu-id="4d90c-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-320">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-321">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-322">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-324">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-324">Office 2016 on Mac</span></span><br><span data-ttu-id="4d90c-325">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4d90c-326">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-326">- TaskPane</span></span><br><span data-ttu-id="4d90c-327">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-327">
        - Content</span></span></td>
    <td><span data-ttu-id="4d90c-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4d90c-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4d90c-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="4d90c-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-331">- BindingEvents</span></span><br><span data-ttu-id="4d90c-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-332">
        - CompressedFile</span></span><br><span data-ttu-id="4d90c-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-333">
        - DocumentEvents</span></span><br><span data-ttu-id="4d90c-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-334">
        - File</span></span><br><span data-ttu-id="4d90c-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-335">
        - MatrixBindings</span></span><br><span data-ttu-id="4d90c-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-337">
        - PdfFile</span></span><br><span data-ttu-id="4d90c-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-338">
        - Selection</span></span><br><span data-ttu-id="4d90c-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-339">
        - Settings</span></span><br><span data-ttu-id="4d90c-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-340">
        - TableBindings</span></span><br><span data-ttu-id="4d90c-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-341">
        - TableCoercion</span></span><br><span data-ttu-id="4d90c-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-342">
        - TextBindings</span></span><br><span data-ttu-id="4d90c-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d90c-344">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d90c-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="4d90c-345">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4d90c-346">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4d90c-347">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4d90c-348">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4d90c-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-350">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-350">Office on the web</span></span></td>
    <td><span data-ttu-id="4d90c-351">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="4d90c-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-353">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-353">Office on Windows</span></span><br><span data-ttu-id="4d90c-354">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="4d90c-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="4d90c-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="4d90c-357">Office for Mac</span></span><br><span data-ttu-id="4d90c-358">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="4d90c-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="4d90c-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="4d90c-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="4d90c-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="4d90c-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="4d90c-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d90c-362">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-362">Platform</span></span></th>
    <th><span data-ttu-id="4d90c-363">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-363">Extension points</span></span></th>
    <th><span data-ttu-id="4d90c-364">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d90c-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-366">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-366">Office on the web</span></span><br><span data-ttu-id="4d90c-367">(モダン)</span><span class="sxs-lookup"><span data-stu-id="4d90c-367">(modern)</span></span></td>
    <td> <span data-ttu-id="4d90c-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-368">- Mail Read</span></span><br><span data-ttu-id="4d90c-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-369">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d90c-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4d90c-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="4d90c-379">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-380">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-380">Office on the web</span></span><br><span data-ttu-id="4d90c-381">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="4d90c-381">(classic)</span></span></td>
    <td> <span data-ttu-id="4d90c-382">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-382">- Mail Read</span></span><br><span data-ttu-id="4d90c-383">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-383">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d90c-391">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-392">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-392">Office on Windows</span></span><br><span data-ttu-id="4d90c-393">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-394">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-394">- Mail Read</span></span><br><span data-ttu-id="4d90c-395">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-395">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d90c-397">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4d90c-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d90c-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d90c-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4d90c-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="4d90c-406">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-407">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-407">Office 2019 on Windows</span></span><br><span data-ttu-id="4d90c-408">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-409">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-409">- Mail Read</span></span><br><span data-ttu-id="4d90c-410">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-410">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d90c-412">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4d90c-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d90c-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d90c-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="4d90c-420">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-421">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-421">Office 2016 on Windows</span></span><br><span data-ttu-id="4d90c-422">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-423">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-423">- Mail Read</span></span><br><span data-ttu-id="4d90c-424">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-424">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4d90c-426">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="4d90c-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="4d90c-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="4d90c-431">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-432">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4d90c-432">Office 2013 on Windows</span></span><br><span data-ttu-id="4d90c-433">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-434">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-434">- Mail Read</span></span><br><span data-ttu-id="4d90c-435">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="4d90c-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="4d90c-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="4d90c-440">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-441">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-441">Office on iOS</span></span><br><span data-ttu-id="4d90c-442">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-443">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-443">- Mail Read</span></span><br><span data-ttu-id="4d90c-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4d90c-450">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-451">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-451">Office on Mac</span></span><br><span data-ttu-id="4d90c-452">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-453">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-453">- Mail Read</span></span><br><span data-ttu-id="4d90c-454">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-454">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4d90c-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4d90c-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="4d90c-464">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-465">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-465">Office 2019 on Mac</span></span><br><span data-ttu-id="4d90c-466">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-467">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-467">- Mail Read</span></span><br><span data-ttu-id="4d90c-468">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-468">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d90c-476">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-477">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-477">Office 2016 on Mac</span></span><br><span data-ttu-id="4d90c-478">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-479">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-479">- Mail Read</span></span><br><span data-ttu-id="4d90c-480">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="4d90c-480">
      - Mail Compose</span></span><br><span data-ttu-id="4d90c-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4d90c-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="4d90c-488">使用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-489">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-489">Office on Android</span></span><br><span data-ttu-id="4d90c-490">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-491">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="4d90c-491">- Mail Read</span></span><br><span data-ttu-id="4d90c-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4d90c-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4d90c-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4d90c-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4d90c-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="4d90c-498">利用不可</span><span class="sxs-lookup"><span data-stu-id="4d90c-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d90c-499">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d90c-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4d90c-500">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="4d90c-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="4d90c-501">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4d90c-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="4d90c-502">Word</span><span class="sxs-lookup"><span data-stu-id="4d90c-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d90c-503">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-503">Platform</span></span></th>
    <th><span data-ttu-id="4d90c-504">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-504">Extension points</span></span></th>
    <th><span data-ttu-id="4d90c-505">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d90c-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-507">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="4d90c-508">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-508">- TaskPane</span></span><br><span data-ttu-id="4d90c-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4d90c-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4d90c-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4d90c-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-516">- BindingEvents</span></span><br><span data-ttu-id="4d90c-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-518">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-519">
         - File</span></span><br><span data-ttu-id="4d90c-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-521">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-524">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-525">
         - Selection</span></span><br><span data-ttu-id="4d90c-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-526">
         - Settings</span></span><br><span data-ttu-id="4d90c-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-527">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-528">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-529">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-530">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-532">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-532">Office on Windows</span></span><br><span data-ttu-id="4d90c-533">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-534">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-534">- TaskPane</span></span><br><span data-ttu-id="4d90c-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4d90c-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4d90c-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4d90c-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-542">- BindingEvents</span></span><br><span data-ttu-id="4d90c-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-543">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-545">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-546">
         - File</span></span><br><span data-ttu-id="4d90c-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-548">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-551">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-552">
         - Selection</span></span><br><span data-ttu-id="4d90c-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-553">
         - Settings</span></span><br><span data-ttu-id="4d90c-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-554">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-555">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-556">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-557">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-559">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-559">Office 2019 on Windows</span></span><br><span data-ttu-id="4d90c-560">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-561">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-561">- TaskPane</span></span><br><span data-ttu-id="4d90c-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4d90c-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4d90c-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-568">- BindingEvents</span></span><br><span data-ttu-id="4d90c-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-569">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-571">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-572">
         - File</span></span><br><span data-ttu-id="4d90c-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-574">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-577">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-578">
         - Selection</span></span><br><span data-ttu-id="4d90c-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-579">
         - Settings</span></span><br><span data-ttu-id="4d90c-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-580">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-581">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-582">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-583">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-585">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-585">Office 2016 on Windows</span></span><br><span data-ttu-id="4d90c-586">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-587">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4d90c-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-591">- BindingEvents</span></span><br><span data-ttu-id="4d90c-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-592">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-594">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-595">
         - File</span></span><br><span data-ttu-id="4d90c-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-597">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-600">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-601">
         - Selection</span></span><br><span data-ttu-id="4d90c-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-602">
         - Settings</span></span><br><span data-ttu-id="4d90c-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-603">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-604">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-605">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-606">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-608">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4d90c-608">Office 2013 on Windows</span></span><br><span data-ttu-id="4d90c-609">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-610">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d90c-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4d90c-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-613">- BindingEvents</span></span><br><span data-ttu-id="4d90c-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-614">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-616">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-617">
         - File</span></span><br><span data-ttu-id="4d90c-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-619">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-622">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-623">
         - Selection</span></span><br><span data-ttu-id="4d90c-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-624">
         - Settings</span></span><br><span data-ttu-id="4d90c-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-625">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-626">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-627">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-628">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-630">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-630">Office on iPad</span></span><br><span data-ttu-id="4d90c-631">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-632">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4d90c-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4d90c-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="4d90c-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-638">- BindingEvents</span></span><br><span data-ttu-id="4d90c-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-639">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-641">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-642">
         - File</span></span><br><span data-ttu-id="4d90c-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-644">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-647">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-648">
         - Selection</span></span><br><span data-ttu-id="4d90c-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-649">
         - Settings</span></span><br><span data-ttu-id="4d90c-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-650">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-651">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-652">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-653">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-655">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-655">Office on Mac</span></span><br><span data-ttu-id="4d90c-656">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-657">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-657">- TaskPane</span></span><br><span data-ttu-id="4d90c-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4d90c-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4d90c-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="4d90c-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-665">- BindingEvents</span></span><br><span data-ttu-id="4d90c-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-666">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-668">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-669">
         - File</span></span><br><span data-ttu-id="4d90c-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-671">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-674">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-675">
         - Selection</span></span><br><span data-ttu-id="4d90c-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-676">
         - Settings</span></span><br><span data-ttu-id="4d90c-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-677">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-678">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-679">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-680">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-682">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-682">Office 2019 on Mac</span></span><br><span data-ttu-id="4d90c-683">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-684">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-684">- TaskPane</span></span><br><span data-ttu-id="4d90c-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4d90c-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4d90c-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="4d90c-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-691">- BindingEvents</span></span><br><span data-ttu-id="4d90c-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-692">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-694">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-695">
         - File</span></span><br><span data-ttu-id="4d90c-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-697">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-700">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-701">
         - Selection</span></span><br><span data-ttu-id="4d90c-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-702">
         - Settings</span></span><br><span data-ttu-id="4d90c-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-703">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-704">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-705">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-706">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-708">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-708">Office 2016 on Mac</span></span><br><span data-ttu-id="4d90c-709">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-710">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4d90c-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4d90c-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4d90c-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-714">- BindingEvents</span></span><br><span data-ttu-id="4d90c-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-715">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="4d90c-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="4d90c-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-717">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-718">
         - File</span></span><br><span data-ttu-id="4d90c-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-720">
         - MatrixBindings</span></span><br><span data-ttu-id="4d90c-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="4d90c-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="4d90c-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-723">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-724">
         - Selection</span></span><br><span data-ttu-id="4d90c-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-725">
         - Settings</span></span><br><span data-ttu-id="4d90c-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-726">
         - TableBindings</span></span><br><span data-ttu-id="4d90c-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-727">
         - TableCoercion</span></span><br><span data-ttu-id="4d90c-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="4d90c-728">
         - TextBindings</span></span><br><span data-ttu-id="4d90c-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-729">
         - TextCoercion</span></span><br><span data-ttu-id="4d90c-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="4d90c-731">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d90c-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4d90c-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4d90c-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d90c-733">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-733">Platform</span></span></th>
    <th><span data-ttu-id="4d90c-734">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-734">Extension points</span></span></th>
    <th><span data-ttu-id="4d90c-735">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d90c-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-737">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="4d90c-738">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-738">- Content</span></span><br><span data-ttu-id="4d90c-739">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-739">
         - TaskPane</span></span><br><span data-ttu-id="4d90c-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4d90c-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4d90c-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-745">- ActiveView</span></span><br><span data-ttu-id="4d90c-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-746">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-747">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-748">
         - File</span></span><br><span data-ttu-id="4d90c-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-749">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-750">
         - Selection</span></span><br><span data-ttu-id="4d90c-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-751">
         - Settings</span></span><br><span data-ttu-id="4d90c-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-753">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-753">Office on Windows</span></span><br><span data-ttu-id="4d90c-754">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-755">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-755">- Content</span></span><br><span data-ttu-id="4d90c-756">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-756">
         - TaskPane</span></span><br><span data-ttu-id="4d90c-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4d90c-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4d90c-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-762">- ActiveView</span></span><br><span data-ttu-id="4d90c-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-763">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-764">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-765">
         - File</span></span><br><span data-ttu-id="4d90c-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-766">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-767">
         - Selection</span></span><br><span data-ttu-id="4d90c-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-768">
         - Settings</span></span><br><span data-ttu-id="4d90c-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-770">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-770">Office 2019 on Windows</span></span><br><span data-ttu-id="4d90c-771">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-772">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-772">- Content</span></span><br><span data-ttu-id="4d90c-773">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-773">
         - TaskPane</span></span><br><span data-ttu-id="4d90c-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-777">- ActiveView</span></span><br><span data-ttu-id="4d90c-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-778">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-779">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-780">
         - File</span></span><br><span data-ttu-id="4d90c-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-781">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-782">
         - Selection</span></span><br><span data-ttu-id="4d90c-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-783">
         - Settings</span></span><br><span data-ttu-id="4d90c-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-785">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-785">Office 2016 on Windows</span></span><br><span data-ttu-id="4d90c-786">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-787">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-787">- Content</span></span><br><span data-ttu-id="4d90c-788">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d90c-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4d90c-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-791">- ActiveView</span></span><br><span data-ttu-id="4d90c-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-792">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-793">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-794">
         - File</span></span><br><span data-ttu-id="4d90c-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-795">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-796">
         - Selection</span></span><br><span data-ttu-id="4d90c-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-797">
         - Settings</span></span><br><span data-ttu-id="4d90c-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-799">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4d90c-799">Office 2013 on Windows</span></span><br><span data-ttu-id="4d90c-800">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-801">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-801">- Content</span></span><br><span data-ttu-id="4d90c-802">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="4d90c-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d90c-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4d90c-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-805">- ActiveView</span></span><br><span data-ttu-id="4d90c-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-806">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-807">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-808">
         - File</span></span><br><span data-ttu-id="4d90c-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-809">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-810">
         - Selection</span></span><br><span data-ttu-id="4d90c-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-811">
         - Settings</span></span><br><span data-ttu-id="4d90c-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-813">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-813">Office on iPad</span></span><br><span data-ttu-id="4d90c-814">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-815">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-815">- Content</span></span><br><span data-ttu-id="4d90c-816">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4d90c-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-820">- ActiveView</span></span><br><span data-ttu-id="4d90c-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-821">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-822">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-823">
         - File</span></span><br><span data-ttu-id="4d90c-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-824">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-825">
         - Selection</span></span><br><span data-ttu-id="4d90c-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-826">
         - Settings</span></span><br><span data-ttu-id="4d90c-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-828">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="4d90c-828">Office on Mac</span></span><br><span data-ttu-id="4d90c-829">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="4d90c-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="4d90c-830">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-830">- Content</span></span><br><span data-ttu-id="4d90c-831">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-831">
         - TaskPane</span></span><br><span data-ttu-id="4d90c-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4d90c-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4d90c-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="4d90c-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-837">- ActiveView</span></span><br><span data-ttu-id="4d90c-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-838">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-839">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-840">
         - File</span></span><br><span data-ttu-id="4d90c-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-841">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-842">
         - Selection</span></span><br><span data-ttu-id="4d90c-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-843">
         - Settings</span></span><br><span data-ttu-id="4d90c-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-845">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-845">Office 2019 on Mac</span></span><br><span data-ttu-id="4d90c-846">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-847">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-847">- Content</span></span><br><span data-ttu-id="4d90c-848">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-848">
         - TaskPane</span></span><br><span data-ttu-id="4d90c-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-852">- ActiveView</span></span><br><span data-ttu-id="4d90c-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-853">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-854">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-855">
         - File</span></span><br><span data-ttu-id="4d90c-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-856">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-857">
         - Selection</span></span><br><span data-ttu-id="4d90c-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-858">
         - Settings</span></span><br><span data-ttu-id="4d90c-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-860">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-860">Office 2016 on Mac</span></span><br><span data-ttu-id="4d90c-861">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-862">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-862">- Content</span></span><br><span data-ttu-id="4d90c-863">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="4d90c-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="4d90c-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="4d90c-866">- ActiveView</span></span><br><span data-ttu-id="4d90c-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-867">
         - CompressedFile</span></span><br><span data-ttu-id="4d90c-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-868">
         - DocumentEvents</span></span><br><span data-ttu-id="4d90c-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="4d90c-869">
         - File</span></span><br><span data-ttu-id="4d90c-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="4d90c-870">
         - PdfFile</span></span><br><span data-ttu-id="4d90c-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-871">
         - Selection</span></span><br><span data-ttu-id="4d90c-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-872">
         - Settings</span></span><br><span data-ttu-id="4d90c-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="4d90c-874">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4d90c-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4d90c-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="4d90c-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d90c-876">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-876">Platform</span></span></th>
    <th><span data-ttu-id="4d90c-877">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-877">Extension points</span></span></th>
    <th><span data-ttu-id="4d90c-878">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d90c-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-880">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4d90c-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="4d90c-881">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4d90c-881">- Content</span></span><br><span data-ttu-id="4d90c-882">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-882">
         - TaskPane</span></span><br><span data-ttu-id="4d90c-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="4d90c-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4d90c-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4d90c-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="4d90c-887">- DocumentEvents</span></span><br><span data-ttu-id="4d90c-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="4d90c-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="4d90c-889">
         - Settings</span></span><br><span data-ttu-id="4d90c-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4d90c-891">Project</span><span class="sxs-lookup"><span data-stu-id="4d90c-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4d90c-892">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4d90c-892">Platform</span></span></th>
    <th><span data-ttu-id="4d90c-893">拡張点</span><span class="sxs-lookup"><span data-stu-id="4d90c-893">Extension points</span></span></th>
    <th><span data-ttu-id="4d90c-894">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="4d90c-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4d90c-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-896">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4d90c-896">Office 2019 on Windows</span></span><br><span data-ttu-id="4d90c-897">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-898">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-900">- Selection</span></span><br><span data-ttu-id="4d90c-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-902">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4d90c-902">Office 2016 on Windows</span></span><br><span data-ttu-id="4d90c-903">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-904">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-906">- Selection</span></span><br><span data-ttu-id="4d90c-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4d90c-908">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4d90c-908">Office 2013 on Windows</span></span><br><span data-ttu-id="4d90c-909">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4d90c-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="4d90c-910">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4d90c-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="4d90c-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4d90c-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="4d90c-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="4d90c-912">- Selection</span></span><br><span data-ttu-id="4d90c-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="4d90c-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4d90c-914">関連項目</span><span class="sxs-lookup"><span data-stu-id="4d90c-914">See also</span></span>

- [<span data-ttu-id="4d90c-915">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="4d90c-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4d90c-916">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="4d90c-917">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="4d90c-918">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="4d90c-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="4d90c-919">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="4d90c-919">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="4d90c-920">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="4d90c-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="4d90c-921">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="4d90c-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="4d90c-922">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="4d90c-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="4d90c-923">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="4d90c-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="4d90c-924">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="4d90c-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="4d90c-925">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="4d90c-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
