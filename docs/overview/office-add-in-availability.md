---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 04/13/2020
localization_priority: Priority
ms.openlocfilehash: 72da8db755fe6d1d166f66a70c8c298e5a27adff
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241057"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="72254-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="72254-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="72254-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="72254-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="72254-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="72254-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="72254-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="72254-108">Excel</span><span class="sxs-lookup"><span data-stu-id="72254-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="72254-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="72254-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="72254-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="72254-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="72254-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-114">- TaskPane</span></span><br><span data-ttu-id="72254-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-115">
        - Content</span></span><br><span data-ttu-id="72254-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="72254-116">
        - Custom Functions</span></span><br><span data-ttu-id="72254-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="72254-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="72254-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72254-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72254-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72254-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72254-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72254-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72254-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72254-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="72254-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="72254-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="72254-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="72254-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="72254-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="72254-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-130">
        - BindingEvents</span></span><br><span data-ttu-id="72254-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-131">
        - CompressedFile</span></span><br><span data-ttu-id="72254-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-132">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-133">
        - File</span></span><br><span data-ttu-id="72254-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-134">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-136">
        - Selection</span></span><br><span data-ttu-id="72254-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-137">
        - Settings</span></span><br><span data-ttu-id="72254-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-138">
        - TableBindings</span></span><br><span data-ttu-id="72254-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-139">
        - TableCoercion</span></span><br><span data-ttu-id="72254-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-140">
        - TextBindings</span></span><br><span data-ttu-id="72254-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="72254-142">Office on Windows</span></span><br><span data-ttu-id="72254-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-144">- TaskPane</span></span><br><span data-ttu-id="72254-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-145">
        - Content</span></span><br><span data-ttu-id="72254-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="72254-146">
        - Custom Functions</span></span><br><span data-ttu-id="72254-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="72254-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="72254-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72254-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72254-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72254-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72254-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72254-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72254-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72254-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="72254-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="72254-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="72254-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="72254-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="72254-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-161">
        - BindingEvents</span></span><br><span data-ttu-id="72254-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-162">
        - CompressedFile</span></span><br><span data-ttu-id="72254-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-163">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-164">
        - File</span></span><br><span data-ttu-id="72254-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-165">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-167">
        - Selection</span></span><br><span data-ttu-id="72254-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-168">
        - Settings</span></span><br><span data-ttu-id="72254-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-169">
        - TableBindings</span></span><br><span data-ttu-id="72254-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-170">
        - TableCoercion</span></span><br><span data-ttu-id="72254-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-171">
        - TextBindings</span></span><br><span data-ttu-id="72254-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-173">Office 2019 on Windows</span></span><br><span data-ttu-id="72254-174">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="72254-175">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-175">- TaskPane</span></span><br><span data-ttu-id="72254-176">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-176">
        - Content</span></span><br><span data-ttu-id="72254-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72254-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72254-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72254-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72254-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72254-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72254-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72254-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72254-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-188">- BindingEvents</span></span><br><span data-ttu-id="72254-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-189">
        - CompressedFile</span></span><br><span data-ttu-id="72254-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-190">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-191">
        - File</span></span><br><span data-ttu-id="72254-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-192">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-194">
        - Selection</span></span><br><span data-ttu-id="72254-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-195">
        - Settings</span></span><br><span data-ttu-id="72254-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-196">
        - TableBindings</span></span><br><span data-ttu-id="72254-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-197">
        - TableCoercion</span></span><br><span data-ttu-id="72254-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-198">
        - TextBindings</span></span><br><span data-ttu-id="72254-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-200">Office 2016 on Windows</span></span><br><span data-ttu-id="72254-201">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="72254-202">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-202">- TaskPane</span></span><br><span data-ttu-id="72254-203">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-203">
        - Content</span></span></td>
    <td><span data-ttu-id="72254-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="72254-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-207">- BindingEvents</span></span><br><span data-ttu-id="72254-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-208">
        - CompressedFile</span></span><br><span data-ttu-id="72254-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-209">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-210">
        - File</span></span><br><span data-ttu-id="72254-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-211">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-213">
        - Selection</span></span><br><span data-ttu-id="72254-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-214">
        - Settings</span></span><br><span data-ttu-id="72254-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-215">
        - TableBindings</span></span><br><span data-ttu-id="72254-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-216">
        - TableCoercion</span></span><br><span data-ttu-id="72254-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-217">
        - TextBindings</span></span><br><span data-ttu-id="72254-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="72254-219">Office 2013 on Windows</span></span><br><span data-ttu-id="72254-220">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="72254-221">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-221">
        - TaskPane</span></span><br><span data-ttu-id="72254-222">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="72254-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72254-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="72254-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-225">
        - BindingEvents</span></span><br><span data-ttu-id="72254-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-226">
        - CompressedFile</span></span><br><span data-ttu-id="72254-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-227">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-228">
        - File</span></span><br><span data-ttu-id="72254-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-229">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-231">
        - Selection</span></span><br><span data-ttu-id="72254-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-232">
        - Settings</span></span><br><span data-ttu-id="72254-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-233">
        - TableBindings</span></span><br><span data-ttu-id="72254-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-234">
        - TableCoercion</span></span><br><span data-ttu-id="72254-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-235">
        - TextBindings</span></span><br><span data-ttu-id="72254-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-237">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-237">Office on iPad</span></span><br><span data-ttu-id="72254-238">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="72254-239">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-239">- TaskPane</span></span><br><span data-ttu-id="72254-240">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-240">
        - Content</span></span></td>
    <td><span data-ttu-id="72254-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72254-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72254-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72254-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72254-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72254-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72254-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72254-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="72254-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="72254-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="72254-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="72254-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-253">- BindingEvents</span></span><br><span data-ttu-id="72254-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-254">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-255">
        - File</span></span><br><span data-ttu-id="72254-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-256">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-258">
        - Selection</span></span><br><span data-ttu-id="72254-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-259">
        - Settings</span></span><br><span data-ttu-id="72254-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-260">
        - TableBindings</span></span><br><span data-ttu-id="72254-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-261">
        - TableCoercion</span></span><br><span data-ttu-id="72254-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-262">
        - TextBindings</span></span><br><span data-ttu-id="72254-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-264">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-264">Office on Mac</span></span><br><span data-ttu-id="72254-265">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="72254-266">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-266">- TaskPane</span></span><br><span data-ttu-id="72254-267">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-267">
        - Content</span></span><br><span data-ttu-id="72254-268">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="72254-268">
        - Custom Functions</span></span><br><span data-ttu-id="72254-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72254-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72254-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72254-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72254-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72254-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72254-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72254-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72254-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="72254-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="72254-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="72254-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="72254-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="72254-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-283">- BindingEvents</span></span><br><span data-ttu-id="72254-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-284">
        - CompressedFile</span></span><br><span data-ttu-id="72254-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-285">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-286">
        - File</span></span><br><span data-ttu-id="72254-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-287">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-289">
        - PdfFile</span></span><br><span data-ttu-id="72254-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-290">
        - Selection</span></span><br><span data-ttu-id="72254-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-291">
        - Settings</span></span><br><span data-ttu-id="72254-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-292">
        - TableBindings</span></span><br><span data-ttu-id="72254-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-293">
        - TableCoercion</span></span><br><span data-ttu-id="72254-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-294">
        - TextBindings</span></span><br><span data-ttu-id="72254-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-296">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-296">Office 2019 on Mac</span></span><br><span data-ttu-id="72254-297">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="72254-298">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-298">- TaskPane</span></span><br><span data-ttu-id="72254-299">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-299">
        - Content</span></span><br><span data-ttu-id="72254-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="72254-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="72254-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="72254-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="72254-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="72254-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="72254-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="72254-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="72254-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-311">- BindingEvents</span></span><br><span data-ttu-id="72254-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-312">
        - CompressedFile</span></span><br><span data-ttu-id="72254-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-313">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-314">
        - File</span></span><br><span data-ttu-id="72254-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-315">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-317">
        - PdfFile</span></span><br><span data-ttu-id="72254-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-318">
        - Selection</span></span><br><span data-ttu-id="72254-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-319">
        - Settings</span></span><br><span data-ttu-id="72254-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-320">
        - TableBindings</span></span><br><span data-ttu-id="72254-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-321">
        - TableCoercion</span></span><br><span data-ttu-id="72254-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-322">
        - TextBindings</span></span><br><span data-ttu-id="72254-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-324">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-324">Office 2016 on Mac</span></span><br><span data-ttu-id="72254-325">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="72254-326">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-326">- TaskPane</span></span><br><span data-ttu-id="72254-327">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-327">
        - Content</span></span></td>
    <td><span data-ttu-id="72254-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="72254-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="72254-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="72254-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-331">- BindingEvents</span></span><br><span data-ttu-id="72254-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-332">
        - CompressedFile</span></span><br><span data-ttu-id="72254-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-333">
        - DocumentEvents</span></span><br><span data-ttu-id="72254-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="72254-334">
        - File</span></span><br><span data-ttu-id="72254-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-335">
        - MatrixBindings</span></span><br><span data-ttu-id="72254-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="72254-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-337">
        - PdfFile</span></span><br><span data-ttu-id="72254-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-338">
        - Selection</span></span><br><span data-ttu-id="72254-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-339">
        - Settings</span></span><br><span data-ttu-id="72254-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-340">
        - TableBindings</span></span><br><span data-ttu-id="72254-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-341">
        - TableCoercion</span></span><br><span data-ttu-id="72254-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-342">
        - TextBindings</span></span><br><span data-ttu-id="72254-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="72254-344">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="72254-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="72254-345">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="72254-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="72254-346">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="72254-347">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="72254-348">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="72254-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-350">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-350">Office on the web</span></span></td>
    <td><span data-ttu-id="72254-351">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="72254-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="72254-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-353">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="72254-353">Office on Windows</span></span><br><span data-ttu-id="72254-354">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="72254-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="72254-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="72254-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="72254-357">Office for Mac</span></span><br><span data-ttu-id="72254-358">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="72254-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="72254-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="72254-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="72254-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="72254-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="72254-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72254-362">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-362">Platform</span></span></th>
    <th><span data-ttu-id="72254-363">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-363">Extension points</span></span></th>
    <th><span data-ttu-id="72254-364">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="72254-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-366">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-366">Office on the web</span></span><br><span data-ttu-id="72254-367">(モダン)</span><span class="sxs-lookup"><span data-stu-id="72254-367">(modern)</span></span></td>
    <td> <span data-ttu-id="72254-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-368">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-369">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-370">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-371">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72254-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="72254-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="72254-381">利用不可</span><span class="sxs-lookup"><span data-stu-id="72254-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-382">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-382">Office on the web</span></span><br><span data-ttu-id="72254-383">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="72254-383">(classic)</span></span></td>
    <td> <span data-ttu-id="72254-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-384">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-387">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72254-395">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-396">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="72254-396">Office on Windows</span></span><br><span data-ttu-id="72254-397">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-398">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-399">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72254-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="72254-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="72254-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72254-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="72254-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="72254-412">利用不可</span><span class="sxs-lookup"><span data-stu-id="72254-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-413">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-413">Office 2019 on Windows</span></span><br><span data-ttu-id="72254-414">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-415">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-418">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72254-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="72254-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="72254-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72254-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="72254-428">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-429">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-429">Office 2016 on Windows</span></span><br><span data-ttu-id="72254-430">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-431">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="72254-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="72254-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="72254-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="72254-441">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-442">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="72254-442">Office 2013 on Windows</span></span><br><span data-ttu-id="72254-443">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-444">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-445">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-446">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-447">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="72254-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="72254-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="72254-452">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-453">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-453">Office on iOS</span></span><br><span data-ttu-id="72254-454">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-455">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="72254-462">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-463">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-463">Office on Mac</span></span><br><span data-ttu-id="72254-464">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-465">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-466">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-467">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-468">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="72254-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="72254-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="72254-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="72254-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="72254-478">利用不可</span><span class="sxs-lookup"><span data-stu-id="72254-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-479">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-479">Office 2019 on Mac</span></span><br><span data-ttu-id="72254-480">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-481">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72254-492">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-493">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-493">Office 2016 on Mac</span></span><br><span data-ttu-id="72254-494">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-495">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="72254-496">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="72254-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="72254-497">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="72254-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="72254-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="72254-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="72254-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="72254-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="72254-506">使用不可</span><span class="sxs-lookup"><span data-stu-id="72254-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-507">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-507">Office on Android</span></span><br><span data-ttu-id="72254-508">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="72254-509">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="72254-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">予定の開催者 (作成): オンライン会議</a> (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="72254-510">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="72254-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-511">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-512">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="72254-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="72254-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="72254-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="72254-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="72254-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="72254-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="72254-517">利用不可</span><span class="sxs-lookup"><span data-stu-id="72254-517">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="72254-518">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="72254-518">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="72254-519">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="72254-519">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="72254-520">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="72254-520">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="72254-521">Word</span><span class="sxs-lookup"><span data-stu-id="72254-521">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72254-522">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-522">Platform</span></span></th>
    <th><span data-ttu-id="72254-523">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-523">Extension points</span></span></th>
    <th><span data-ttu-id="72254-524">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-524">API requirement sets</span></span></th>
    <th><span data-ttu-id="72254-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-525"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-526">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-526">Office on the web</span></span></td>
    <td> <span data-ttu-id="72254-527">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-527">- TaskPane</span></span><br><span data-ttu-id="72254-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="72254-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="72254-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="72254-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-535">- BindingEvents</span></span><br><span data-ttu-id="72254-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-537">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-538">
         - File</span></span><br><span data-ttu-id="72254-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-540">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-543">
         - PdfFile</span></span><br><span data-ttu-id="72254-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-544">
         - Selection</span></span><br><span data-ttu-id="72254-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-545">
         - Settings</span></span><br><span data-ttu-id="72254-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-546">
         - TableBindings</span></span><br><span data-ttu-id="72254-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-547">
         - TableCoercion</span></span><br><span data-ttu-id="72254-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-548">
         - TextBindings</span></span><br><span data-ttu-id="72254-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-549">
         - TextCoercion</span></span><br><span data-ttu-id="72254-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-550">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-551">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="72254-551">Office on Windows</span></span><br><span data-ttu-id="72254-552">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-552">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-553">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-553">- TaskPane</span></span><br><span data-ttu-id="72254-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="72254-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="72254-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="72254-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-561">- BindingEvents</span></span><br><span data-ttu-id="72254-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-562">
         - CompressedFile</span></span><br><span data-ttu-id="72254-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-564">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-565">
         - File</span></span><br><span data-ttu-id="72254-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-567">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-570">
         - PdfFile</span></span><br><span data-ttu-id="72254-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-571">
         - Selection</span></span><br><span data-ttu-id="72254-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-572">
         - Settings</span></span><br><span data-ttu-id="72254-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-573">
         - TableBindings</span></span><br><span data-ttu-id="72254-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-574">
         - TableCoercion</span></span><br><span data-ttu-id="72254-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-575">
         - TextBindings</span></span><br><span data-ttu-id="72254-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-576">
         - TextCoercion</span></span><br><span data-ttu-id="72254-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-578">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-578">Office 2019 on Windows</span></span><br><span data-ttu-id="72254-579">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-580">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-580">- TaskPane</span></span><br><span data-ttu-id="72254-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-582">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="72254-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="72254-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-585">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-586">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-587">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-587">- BindingEvents</span></span><br><span data-ttu-id="72254-588">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-588">
         - CompressedFile</span></span><br><span data-ttu-id="72254-589">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-589">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-590">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-590">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-591">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-591">
         - File</span></span><br><span data-ttu-id="72254-592">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-592">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-593">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-593">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-594">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-594">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-595">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-595">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-596">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-596">
         - PdfFile</span></span><br><span data-ttu-id="72254-597">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-597">
         - Selection</span></span><br><span data-ttu-id="72254-598">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-598">
         - Settings</span></span><br><span data-ttu-id="72254-599">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-599">
         - TableBindings</span></span><br><span data-ttu-id="72254-600">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-600">
         - TableCoercion</span></span><br><span data-ttu-id="72254-601">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-601">
         - TextBindings</span></span><br><span data-ttu-id="72254-602">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-602">
         - TextCoercion</span></span><br><span data-ttu-id="72254-603">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-603">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-604">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-604">Office 2016 on Windows</span></span><br><span data-ttu-id="72254-605">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-605">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-606">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-606">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-607">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-608">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="72254-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-609">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-610">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-610">- BindingEvents</span></span><br><span data-ttu-id="72254-611">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-611">
         - CompressedFile</span></span><br><span data-ttu-id="72254-612">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-612">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-613">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-613">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-614">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-614">
         - File</span></span><br><span data-ttu-id="72254-615">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-615">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-616">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-616">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-617">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-617">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-618">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-618">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-619">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-619">
         - PdfFile</span></span><br><span data-ttu-id="72254-620">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-620">
         - Selection</span></span><br><span data-ttu-id="72254-621">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-621">
         - Settings</span></span><br><span data-ttu-id="72254-622">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-622">
         - TableBindings</span></span><br><span data-ttu-id="72254-623">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-623">
         - TableCoercion</span></span><br><span data-ttu-id="72254-624">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-624">
         - TextBindings</span></span><br><span data-ttu-id="72254-625">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-625">
         - TextCoercion</span></span><br><span data-ttu-id="72254-626">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-626">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-627">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="72254-627">Office 2013 on Windows</span></span><br><span data-ttu-id="72254-628">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-628">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-629">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-629">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72254-630">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="72254-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-631">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-632">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-632">- BindingEvents</span></span><br><span data-ttu-id="72254-633">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-633">
         - CompressedFile</span></span><br><span data-ttu-id="72254-634">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-634">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-635">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-635">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-636">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-636">
         - File</span></span><br><span data-ttu-id="72254-637">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-637">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-638">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-638">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-639">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-639">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-640">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-640">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-641">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-641">
         - PdfFile</span></span><br><span data-ttu-id="72254-642">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-642">
         - Selection</span></span><br><span data-ttu-id="72254-643">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-643">
         - Settings</span></span><br><span data-ttu-id="72254-644">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-644">
         - TableBindings</span></span><br><span data-ttu-id="72254-645">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-645">
         - TableCoercion</span></span><br><span data-ttu-id="72254-646">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-646">
         - TextBindings</span></span><br><span data-ttu-id="72254-647">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-647">
         - TextCoercion</span></span><br><span data-ttu-id="72254-648">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-648">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-649">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-649">Office on iPad</span></span><br><span data-ttu-id="72254-650">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-650">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-651">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-651">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="72254-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="72254-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="72254-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-657">- BindingEvents</span></span><br><span data-ttu-id="72254-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-658">
         - CompressedFile</span></span><br><span data-ttu-id="72254-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-660">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-661">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-661">
         - File</span></span><br><span data-ttu-id="72254-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-663">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-666">
         - PdfFile</span></span><br><span data-ttu-id="72254-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-667">
         - Selection</span></span><br><span data-ttu-id="72254-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-668">
         - Settings</span></span><br><span data-ttu-id="72254-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-669">
         - TableBindings</span></span><br><span data-ttu-id="72254-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-670">
         - TableCoercion</span></span><br><span data-ttu-id="72254-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-671">
         - TextBindings</span></span><br><span data-ttu-id="72254-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-672">
         - TextCoercion</span></span><br><span data-ttu-id="72254-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-674">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-674">Office on Mac</span></span><br><span data-ttu-id="72254-675">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-675">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-676">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-676">- TaskPane</span></span><br><span data-ttu-id="72254-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="72254-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="72254-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="72254-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-684">- BindingEvents</span></span><br><span data-ttu-id="72254-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-685">
         - CompressedFile</span></span><br><span data-ttu-id="72254-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-687">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-688">
         - File</span></span><br><span data-ttu-id="72254-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-690">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-693">
         - PdfFile</span></span><br><span data-ttu-id="72254-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-694">
         - Selection</span></span><br><span data-ttu-id="72254-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-695">
         - Settings</span></span><br><span data-ttu-id="72254-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-696">
         - TableBindings</span></span><br><span data-ttu-id="72254-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-697">
         - TableCoercion</span></span><br><span data-ttu-id="72254-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-698">
         - TextBindings</span></span><br><span data-ttu-id="72254-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-699">
         - TextCoercion</span></span><br><span data-ttu-id="72254-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-701">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-701">Office 2019 on Mac</span></span><br><span data-ttu-id="72254-702">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-703">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-703">- TaskPane</span></span><br><span data-ttu-id="72254-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-704">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-705">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="72254-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="72254-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="72254-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="72254-710">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-710">- BindingEvents</span></span><br><span data-ttu-id="72254-711">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-711">
         - CompressedFile</span></span><br><span data-ttu-id="72254-712">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-712">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-713">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-713">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-714">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-714">
         - File</span></span><br><span data-ttu-id="72254-715">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-715">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-716">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-716">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-717">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-717">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-718">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-718">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-719">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-719">
         - PdfFile</span></span><br><span data-ttu-id="72254-720">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-720">
         - Selection</span></span><br><span data-ttu-id="72254-721">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-721">
         - Settings</span></span><br><span data-ttu-id="72254-722">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-722">
         - TableBindings</span></span><br><span data-ttu-id="72254-723">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-723">
         - TableCoercion</span></span><br><span data-ttu-id="72254-724">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-724">
         - TextBindings</span></span><br><span data-ttu-id="72254-725">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-725">
         - TextCoercion</span></span><br><span data-ttu-id="72254-726">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-726">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-727">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-727">Office 2016 on Mac</span></span><br><span data-ttu-id="72254-728">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-728">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-729">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-729">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-730">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="72254-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="72254-731">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="72254-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-732">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-733">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="72254-733">- BindingEvents</span></span><br><span data-ttu-id="72254-734">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-734">
         - CompressedFile</span></span><br><span data-ttu-id="72254-735">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="72254-735">
         - CustomXmlParts</span></span><br><span data-ttu-id="72254-736">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-736">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-737">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-737">
         - File</span></span><br><span data-ttu-id="72254-738">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-738">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-739">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="72254-739">
         - MatrixBindings</span></span><br><span data-ttu-id="72254-740">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-740">
         - MatrixCoercion</span></span><br><span data-ttu-id="72254-741">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-741">
         - OoxmlCoercion</span></span><br><span data-ttu-id="72254-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-742">
         - PdfFile</span></span><br><span data-ttu-id="72254-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-743">
         - Selection</span></span><br><span data-ttu-id="72254-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-744">
         - Settings</span></span><br><span data-ttu-id="72254-745">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="72254-745">
         - TableBindings</span></span><br><span data-ttu-id="72254-746">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-746">
         - TableCoercion</span></span><br><span data-ttu-id="72254-747">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="72254-747">
         - TextBindings</span></span><br><span data-ttu-id="72254-748">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-748">
         - TextCoercion</span></span><br><span data-ttu-id="72254-749">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="72254-749">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="72254-750">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="72254-750">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="72254-751">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="72254-751">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72254-752">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-752">Platform</span></span></th>
    <th><span data-ttu-id="72254-753">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-753">Extension points</span></span></th>
    <th><span data-ttu-id="72254-754">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-754">API requirement sets</span></span></th>
    <th><span data-ttu-id="72254-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-755"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-756">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-756">Office on the web</span></span></td>
    <td> <span data-ttu-id="72254-757">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-757">- Content</span></span><br><span data-ttu-id="72254-758">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-758">
         - TaskPane</span></span><br><span data-ttu-id="72254-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-760">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="72254-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="72254-764">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-764">- ActiveView</span></span><br><span data-ttu-id="72254-765">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-765">
         - CompressedFile</span></span><br><span data-ttu-id="72254-766">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-766">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-767">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-767">
         - File</span></span><br><span data-ttu-id="72254-768">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-768">
         - PdfFile</span></span><br><span data-ttu-id="72254-769">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-769">
         - Selection</span></span><br><span data-ttu-id="72254-770">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-770">
         - Settings</span></span><br><span data-ttu-id="72254-771">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-771">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-772">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="72254-772">Office on Windows</span></span><br><span data-ttu-id="72254-773">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-773">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-774">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-774">- Content</span></span><br><span data-ttu-id="72254-775">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-775">
         - TaskPane</span></span><br><span data-ttu-id="72254-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-777">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="72254-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="72254-781">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-781">- ActiveView</span></span><br><span data-ttu-id="72254-782">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-782">
         - CompressedFile</span></span><br><span data-ttu-id="72254-783">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-783">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-784">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-784">
         - File</span></span><br><span data-ttu-id="72254-785">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-785">
         - PdfFile</span></span><br><span data-ttu-id="72254-786">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-786">
         - Selection</span></span><br><span data-ttu-id="72254-787">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-787">
         - Settings</span></span><br><span data-ttu-id="72254-788">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-788">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-789">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-789">Office 2019 on Windows</span></span><br><span data-ttu-id="72254-790">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-790">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-791">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-791">- Content</span></span><br><span data-ttu-id="72254-792">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-792">
         - TaskPane</span></span><br><span data-ttu-id="72254-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-793">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-796">- ActiveView</span></span><br><span data-ttu-id="72254-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-797">
         - CompressedFile</span></span><br><span data-ttu-id="72254-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-798">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-799">
         - File</span></span><br><span data-ttu-id="72254-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-800">
         - PdfFile</span></span><br><span data-ttu-id="72254-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-801">
         - Selection</span></span><br><span data-ttu-id="72254-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-802">
         - Settings</span></span><br><span data-ttu-id="72254-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-804">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-804">Office 2016 on Windows</span></span><br><span data-ttu-id="72254-805">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-805">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-806">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-806">- Content</span></span><br><span data-ttu-id="72254-807">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72254-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="72254-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-810">- ActiveView</span></span><br><span data-ttu-id="72254-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-811">
         - CompressedFile</span></span><br><span data-ttu-id="72254-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-812">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-813">
         - File</span></span><br><span data-ttu-id="72254-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-814">
         - PdfFile</span></span><br><span data-ttu-id="72254-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-815">
         - Selection</span></span><br><span data-ttu-id="72254-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-816">
         - Settings</span></span><br><span data-ttu-id="72254-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-818">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="72254-818">Office 2013 on Windows</span></span><br><span data-ttu-id="72254-819">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-819">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-820">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-820">- Content</span></span><br><span data-ttu-id="72254-821">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-821">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="72254-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72254-822">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="72254-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-824">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-824">- ActiveView</span></span><br><span data-ttu-id="72254-825">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-825">
         - CompressedFile</span></span><br><span data-ttu-id="72254-826">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-826">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-827">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-827">
         - File</span></span><br><span data-ttu-id="72254-828">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-828">
         - PdfFile</span></span><br><span data-ttu-id="72254-829">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-829">
         - Selection</span></span><br><span data-ttu-id="72254-830">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-830">
         - Settings</span></span><br><span data-ttu-id="72254-831">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-831">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-832">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-832">Office on iPad</span></span><br><span data-ttu-id="72254-833">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-833">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-834">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-834">- Content</span></span><br><span data-ttu-id="72254-835">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-835">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-836">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="72254-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-839">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-839">- ActiveView</span></span><br><span data-ttu-id="72254-840">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-840">
         - CompressedFile</span></span><br><span data-ttu-id="72254-841">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-841">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-842">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-842">
         - File</span></span><br><span data-ttu-id="72254-843">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-843">
         - PdfFile</span></span><br><span data-ttu-id="72254-844">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-844">
         - Selection</span></span><br><span data-ttu-id="72254-845">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-845">
         - Settings</span></span><br><span data-ttu-id="72254-846">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-846">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-847">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="72254-847">Office on Mac</span></span><br><span data-ttu-id="72254-848">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="72254-848">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="72254-849">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-849">- Content</span></span><br><span data-ttu-id="72254-850">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-850">
         - TaskPane</span></span><br><span data-ttu-id="72254-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-852">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="72254-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="72254-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="72254-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="72254-856">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-856">- ActiveView</span></span><br><span data-ttu-id="72254-857">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-857">
         - CompressedFile</span></span><br><span data-ttu-id="72254-858">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-858">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-859">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-859">
         - File</span></span><br><span data-ttu-id="72254-860">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-860">
         - PdfFile</span></span><br><span data-ttu-id="72254-861">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-861">
         - Selection</span></span><br><span data-ttu-id="72254-862">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-862">
         - Settings</span></span><br><span data-ttu-id="72254-863">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-863">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-864">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-864">Office 2019 on Mac</span></span><br><span data-ttu-id="72254-865">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-865">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-866">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-866">- Content</span></span><br><span data-ttu-id="72254-867">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-867">
         - TaskPane</span></span><br><span data-ttu-id="72254-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-868">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-869">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-870">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-871">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-871">- ActiveView</span></span><br><span data-ttu-id="72254-872">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-872">
         - CompressedFile</span></span><br><span data-ttu-id="72254-873">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-873">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-874">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-874">
         - File</span></span><br><span data-ttu-id="72254-875">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-875">
         - PdfFile</span></span><br><span data-ttu-id="72254-876">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-876">
         - Selection</span></span><br><span data-ttu-id="72254-877">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-877">
         - Settings</span></span><br><span data-ttu-id="72254-878">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-878">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-879">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-879">Office 2016 on Mac</span></span><br><span data-ttu-id="72254-880">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-880">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-881">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-881">- Content</span></span><br><span data-ttu-id="72254-882">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-882">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="72254-883">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="72254-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-884">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-885">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="72254-885">- ActiveView</span></span><br><span data-ttu-id="72254-886">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="72254-886">
         - CompressedFile</span></span><br><span data-ttu-id="72254-887">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-887">
         - DocumentEvents</span></span><br><span data-ttu-id="72254-888">
         - File</span><span class="sxs-lookup"><span data-stu-id="72254-888">
         - File</span></span><br><span data-ttu-id="72254-889">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="72254-889">
         - PdfFile</span></span><br><span data-ttu-id="72254-890">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="72254-890">
         - Selection</span></span><br><span data-ttu-id="72254-891">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-891">
         - Settings</span></span><br><span data-ttu-id="72254-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-892">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="72254-893">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="72254-893">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="72254-894">OneNote</span><span class="sxs-lookup"><span data-stu-id="72254-894">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72254-895">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-895">Platform</span></span></th>
    <th><span data-ttu-id="72254-896">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-896">Extension points</span></span></th>
    <th><span data-ttu-id="72254-897">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-897">API requirement sets</span></span></th>
    <th><span data-ttu-id="72254-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-898"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-899">Office on the web</span><span class="sxs-lookup"><span data-stu-id="72254-899">Office on the web</span></span></td>
    <td> <span data-ttu-id="72254-900">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="72254-900">- Content</span></span><br><span data-ttu-id="72254-901">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-901">
         - TaskPane</span></span><br><span data-ttu-id="72254-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="72254-902">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="72254-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-903">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="72254-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="72254-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-906">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="72254-906">- DocumentEvents</span></span><br><span data-ttu-id="72254-907">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-907">
         - HtmlCoercion</span></span><br><span data-ttu-id="72254-908">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="72254-908">
         - Settings</span></span><br><span data-ttu-id="72254-909">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-909">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="72254-910">Project</span><span class="sxs-lookup"><span data-stu-id="72254-910">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="72254-911">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="72254-911">Platform</span></span></th>
    <th><span data-ttu-id="72254-912">拡張点</span><span class="sxs-lookup"><span data-stu-id="72254-912">Extension points</span></span></th>
    <th><span data-ttu-id="72254-913">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-913">API requirement sets</span></span></th>
    <th><span data-ttu-id="72254-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="72254-914"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-915">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="72254-915">Office 2019 on Windows</span></span><br><span data-ttu-id="72254-916">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-916">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-917">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-917">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-918">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-919">- Selection</span><span class="sxs-lookup"><span data-stu-id="72254-919">- Selection</span></span><br><span data-ttu-id="72254-920">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-920">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-921">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="72254-921">Office 2016 on Windows</span></span><br><span data-ttu-id="72254-922">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-922">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-923">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-923">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-924">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-925">- Selection</span><span class="sxs-lookup"><span data-stu-id="72254-925">- Selection</span></span><br><span data-ttu-id="72254-926">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-926">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="72254-927">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="72254-927">Office 2013 on Windows</span></span><br><span data-ttu-id="72254-928">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="72254-928">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="72254-929">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="72254-929">- TaskPane</span></span></td>
    <td> <span data-ttu-id="72254-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="72254-930">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="72254-931">- Selection</span><span class="sxs-lookup"><span data-stu-id="72254-931">- Selection</span></span><br><span data-ttu-id="72254-932">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="72254-932">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="72254-933">関連項目</span><span class="sxs-lookup"><span data-stu-id="72254-933">See also</span></span>

- [<span data-ttu-id="72254-934">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="72254-934">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="72254-935">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-935">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="72254-936">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-936">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="72254-937">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="72254-937">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="72254-938">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="72254-938">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="72254-939">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="72254-939">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="72254-940">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="72254-940">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="72254-941">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="72254-941">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="72254-942">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="72254-942">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="72254-943">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="72254-943">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="72254-944">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="72254-944">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="72254-945">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="72254-945">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)