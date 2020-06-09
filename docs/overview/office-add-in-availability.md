---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 8c3c187d8f9b70f40a35e3773a2267dc76decbd0
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611983"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="c806f-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="c806f-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="c806f-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="c806f-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="c806f-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="c806f-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="c806f-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="c806f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="c806f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c806f-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c806f-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c806f-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c806f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="c806f-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-114">- TaskPane</span></span><br><span data-ttu-id="c806f-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-115">
        - Content</span></span><br><span data-ttu-id="c806f-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c806f-116">
        - Custom Functions</span></span><br><span data-ttu-id="c806f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="c806f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c806f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c806f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c806f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c806f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c806f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c806f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c806f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c806f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c806f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c806f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c806f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c806f-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c806f-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c806f-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="c806f-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="c806f-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-131">
        - BindingEvents</span></span><br><span data-ttu-id="c806f-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-132">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-133">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-134">
        - File</span></span><br><span data-ttu-id="c806f-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-135">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-137">
        - Selection</span></span><br><span data-ttu-id="c806f-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-138">
        - Settings</span></span><br><span data-ttu-id="c806f-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-139">
        - TableBindings</span></span><br><span data-ttu-id="c806f-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-140">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-141">
        - TextBindings</span></span><br><span data-ttu-id="c806f-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-143">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="c806f-143">Office on Windows</span></span><br><span data-ttu-id="c806f-144">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-145">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-145">- TaskPane</span></span><br><span data-ttu-id="c806f-146">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-146">
        - Content</span></span><br><span data-ttu-id="c806f-147">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c806f-147">
        - Custom Functions</span></span><br><span data-ttu-id="c806f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="c806f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="c806f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c806f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c806f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c806f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c806f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c806f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c806f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c806f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c806f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c806f-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c806f-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c806f-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c806f-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c806f-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c806f-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-163">
        - BindingEvents</span></span><br><span data-ttu-id="c806f-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-164">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-165">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-166">
        - File</span></span><br><span data-ttu-id="c806f-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-167">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-169">
        - Selection</span></span><br><span data-ttu-id="c806f-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-170">
        - Settings</span></span><br><span data-ttu-id="c806f-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-171">
        - TableBindings</span></span><br><span data-ttu-id="c806f-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-172">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-173">
        - TextBindings</span></span><br><span data-ttu-id="c806f-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-175">Office 2019 on Windows</span></span><br><span data-ttu-id="c806f-176">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c806f-177">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-177">- TaskPane</span></span><br><span data-ttu-id="c806f-178">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-178">
        - Content</span></span><br><span data-ttu-id="c806f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c806f-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c806f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c806f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c806f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c806f-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c806f-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c806f-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c806f-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-190">- BindingEvents</span></span><br><span data-ttu-id="c806f-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-191">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-192">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-193">
        - File</span></span><br><span data-ttu-id="c806f-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-194">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-196">
        - Selection</span></span><br><span data-ttu-id="c806f-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-197">
        - Settings</span></span><br><span data-ttu-id="c806f-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-198">
        - TableBindings</span></span><br><span data-ttu-id="c806f-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-199">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-200">
        - TextBindings</span></span><br><span data-ttu-id="c806f-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-202">Office 2016 on Windows</span></span><br><span data-ttu-id="c806f-203">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c806f-204">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-204">- TaskPane</span></span><br><span data-ttu-id="c806f-205">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-205">
        - Content</span></span></td>
    <td><span data-ttu-id="c806f-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c806f-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-209">- BindingEvents</span></span><br><span data-ttu-id="c806f-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-210">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-211">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-212">
        - File</span></span><br><span data-ttu-id="c806f-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-213">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-215">
        - Selection</span></span><br><span data-ttu-id="c806f-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-216">
        - Settings</span></span><br><span data-ttu-id="c806f-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-217">
        - TableBindings</span></span><br><span data-ttu-id="c806f-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-218">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-219">
        - TextBindings</span></span><br><span data-ttu-id="c806f-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c806f-221">Office 2013 on Windows</span></span><br><span data-ttu-id="c806f-222">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c806f-223">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-223">
        - TaskPane</span></span><br><span data-ttu-id="c806f-224">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="c806f-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c806f-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c806f-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-227">
        - BindingEvents</span></span><br><span data-ttu-id="c806f-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-228">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-229">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-230">
        - File</span></span><br><span data-ttu-id="c806f-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-231">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-233">
        - Selection</span></span><br><span data-ttu-id="c806f-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-234">
        - Settings</span></span><br><span data-ttu-id="c806f-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-235">
        - TableBindings</span></span><br><span data-ttu-id="c806f-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-236">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-237">
        - TextBindings</span></span><br><span data-ttu-id="c806f-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-239">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-239">Office on iPad</span></span><br><span data-ttu-id="c806f-240">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c806f-241">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-241">- TaskPane</span></span><br><span data-ttu-id="c806f-242">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-242">
        - Content</span></span></td>
    <td><span data-ttu-id="c806f-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c806f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c806f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c806f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c806f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c806f-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c806f-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c806f-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c806f-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c806f-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c806f-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c806f-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c806f-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c806f-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-256">- BindingEvents</span></span><br><span data-ttu-id="c806f-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-257">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-258">
        - File</span></span><br><span data-ttu-id="c806f-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-259">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-261">
        - Selection</span></span><br><span data-ttu-id="c806f-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-262">
        - Settings</span></span><br><span data-ttu-id="c806f-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-263">
        - TableBindings</span></span><br><span data-ttu-id="c806f-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-264">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-265">
        - TextBindings</span></span><br><span data-ttu-id="c806f-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-267">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-267">Office on Mac</span></span><br><span data-ttu-id="c806f-268">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c806f-269">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-269">- TaskPane</span></span><br><span data-ttu-id="c806f-270">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-270">
        - Content</span></span><br><span data-ttu-id="c806f-271">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c806f-271">
        - Custom Functions</span></span><br><span data-ttu-id="c806f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c806f-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c806f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c806f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c806f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c806f-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c806f-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c806f-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c806f-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="c806f-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="c806f-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="c806f-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="c806f-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="c806f-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="c806f-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="c806f-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-287">- BindingEvents</span></span><br><span data-ttu-id="c806f-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-288">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-289">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-290">
        - File</span></span><br><span data-ttu-id="c806f-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-291">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-293">
        - PdfFile</span></span><br><span data-ttu-id="c806f-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-294">
        - Selection</span></span><br><span data-ttu-id="c806f-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-295">
        - Settings</span></span><br><span data-ttu-id="c806f-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-296">
        - TableBindings</span></span><br><span data-ttu-id="c806f-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-297">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-298">
        - TextBindings</span></span><br><span data-ttu-id="c806f-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-300">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-300">Office 2019 on Mac</span></span><br><span data-ttu-id="c806f-301">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c806f-302">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-302">- TaskPane</span></span><br><span data-ttu-id="c806f-303">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-303">
        - Content</span></span><br><span data-ttu-id="c806f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="c806f-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="c806f-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="c806f-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="c806f-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="c806f-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="c806f-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="c806f-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="c806f-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-315">- BindingEvents</span></span><br><span data-ttu-id="c806f-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-316">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-317">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-318">
        - File</span></span><br><span data-ttu-id="c806f-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-319">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-321">
        - PdfFile</span></span><br><span data-ttu-id="c806f-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-322">
        - Selection</span></span><br><span data-ttu-id="c806f-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-323">
        - Settings</span></span><br><span data-ttu-id="c806f-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-324">
        - TableBindings</span></span><br><span data-ttu-id="c806f-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-325">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-326">
        - TextBindings</span></span><br><span data-ttu-id="c806f-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-328">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-328">Office 2016 on Mac</span></span><br><span data-ttu-id="c806f-329">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="c806f-330">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-330">- TaskPane</span></span><br><span data-ttu-id="c806f-331">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-331">
        - Content</span></span></td>
    <td><span data-ttu-id="c806f-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="c806f-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c806f-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="c806f-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-335">- BindingEvents</span></span><br><span data-ttu-id="c806f-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-336">
        - CompressedFile</span></span><br><span data-ttu-id="c806f-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-337">
        - DocumentEvents</span></span><br><span data-ttu-id="c806f-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="c806f-338">
        - File</span></span><br><span data-ttu-id="c806f-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-339">
        - MatrixBindings</span></span><br><span data-ttu-id="c806f-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="c806f-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-341">
        - PdfFile</span></span><br><span data-ttu-id="c806f-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-342">
        - Selection</span></span><br><span data-ttu-id="c806f-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-343">
        - Settings</span></span><br><span data-ttu-id="c806f-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-344">
        - TableBindings</span></span><br><span data-ttu-id="c806f-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-345">
        - TableCoercion</span></span><br><span data-ttu-id="c806f-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-346">
        - TextBindings</span></span><br><span data-ttu-id="c806f-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c806f-348">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="c806f-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="c806f-349">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="c806f-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="c806f-350">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="c806f-351">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="c806f-352">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="c806f-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-354">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-354">Office on the web</span></span></td>
    <td><span data-ttu-id="c806f-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c806f-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c806f-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-357">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="c806f-357">Office on Windows</span></span><br><span data-ttu-id="c806f-358">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="c806f-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c806f-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c806f-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-361">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="c806f-361">Office for Mac</span></span><br><span data-ttu-id="c806f-362">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="c806f-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="c806f-363">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="c806f-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="c806f-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="c806f-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="c806f-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c806f-366">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-366">Platform</span></span></th>
    <th><span data-ttu-id="c806f-367">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-367">Extension points</span></span></th>
    <th><span data-ttu-id="c806f-368">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="c806f-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-370">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-370">Office on the web</span></span><br><span data-ttu-id="c806f-371">(モダン)</span><span class="sxs-lookup"><span data-stu-id="c806f-371">(modern)</span></span></td>
    <td> <span data-ttu-id="c806f-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c806f-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c806f-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c806f-385">利用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-386">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-386">Office on the web</span></span><br><span data-ttu-id="c806f-387">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="c806f-387">(classic)</span></span></td>
    <td> <span data-ttu-id="c806f-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c806f-399">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-400">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="c806f-400">Office on Windows</span></span><br><span data-ttu-id="c806f-401">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c806f-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="c806f-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c806f-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c806f-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c806f-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c806f-416">利用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-417">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-417">Office 2019 on Windows</span></span><br><span data-ttu-id="c806f-418">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c806f-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="c806f-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c806f-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c806f-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="c806f-432">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-433">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-433">Office 2016 on Windows</span></span><br><span data-ttu-id="c806f-434">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="c806f-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="c806f-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="c806f-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c806f-445">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-446">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c806f-446">Office 2013 on Windows</span></span><br><span data-ttu-id="c806f-447">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="c806f-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="c806f-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="c806f-456">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-457">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-457">Office on iOS</span></span><br><span data-ttu-id="c806f-458">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c806f-466">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-467">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-467">Office on Mac</span></span><br><span data-ttu-id="c806f-468">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="c806f-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="c806f-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="c806f-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="c806f-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="c806f-482">利用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-483">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-483">Office 2019 on Mac</span></span><br><span data-ttu-id="c806f-484">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c806f-496">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-497">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-497">Office 2016 on Mac</span></span><br><span data-ttu-id="c806f-498">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="c806f-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="c806f-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="c806f-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="c806f-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="c806f-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="c806f-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="c806f-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="c806f-510">使用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-511">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-511">Office on Android</span></span><br><span data-ttu-id="c806f-512">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="c806f-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="c806f-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">予定の開催者 (作成): オンライン会議</a> (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c806f-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="c806f-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="c806f-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="c806f-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="c806f-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="c806f-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="c806f-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="c806f-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="c806f-521">利用不可</span><span class="sxs-lookup"><span data-stu-id="c806f-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="c806f-522">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="c806f-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c806f-523">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="c806f-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="c806f-524">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c806f-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="c806f-525">Word</span><span class="sxs-lookup"><span data-stu-id="c806f-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c806f-526">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-526">Platform</span></span></th>
    <th><span data-ttu-id="c806f-527">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-527">Extension points</span></span></th>
    <th><span data-ttu-id="c806f-528">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="c806f-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-530">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="c806f-531">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-531">- TaskPane</span></span><br><span data-ttu-id="c806f-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c806f-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c806f-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c806f-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-539">- BindingEvents</span></span><br><span data-ttu-id="c806f-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-541">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-542">
         - File</span></span><br><span data-ttu-id="c806f-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-544">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-547">
         - PdfFile</span></span><br><span data-ttu-id="c806f-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-548">
         - Selection</span></span><br><span data-ttu-id="c806f-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-549">
         - Settings</span></span><br><span data-ttu-id="c806f-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-550">
         - TableBindings</span></span><br><span data-ttu-id="c806f-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-551">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-552">
         - TextBindings</span></span><br><span data-ttu-id="c806f-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-553">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-555">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="c806f-555">Office on Windows</span></span><br><span data-ttu-id="c806f-556">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-557">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-557">- TaskPane</span></span><br><span data-ttu-id="c806f-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c806f-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c806f-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c806f-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-565">- BindingEvents</span></span><br><span data-ttu-id="c806f-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-566">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-568">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-569">
         - File</span></span><br><span data-ttu-id="c806f-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-571">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-574">
         - PdfFile</span></span><br><span data-ttu-id="c806f-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-575">
         - Selection</span></span><br><span data-ttu-id="c806f-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-576">
         - Settings</span></span><br><span data-ttu-id="c806f-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-577">
         - TableBindings</span></span><br><span data-ttu-id="c806f-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-578">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-579">
         - TextBindings</span></span><br><span data-ttu-id="c806f-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-580">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-582">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-582">Office 2019 on Windows</span></span><br><span data-ttu-id="c806f-583">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-584">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-584">- TaskPane</span></span><br><span data-ttu-id="c806f-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c806f-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c806f-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-591">- BindingEvents</span></span><br><span data-ttu-id="c806f-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-592">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-594">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-595">
         - File</span></span><br><span data-ttu-id="c806f-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-597">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-600">
         - PdfFile</span></span><br><span data-ttu-id="c806f-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-601">
         - Selection</span></span><br><span data-ttu-id="c806f-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-602">
         - Settings</span></span><br><span data-ttu-id="c806f-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-603">
         - TableBindings</span></span><br><span data-ttu-id="c806f-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-604">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-605">
         - TextBindings</span></span><br><span data-ttu-id="c806f-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-606">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-608">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-608">Office 2016 on Windows</span></span><br><span data-ttu-id="c806f-609">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-610">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c806f-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-614">- BindingEvents</span></span><br><span data-ttu-id="c806f-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-615">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-617">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-618">
         - File</span></span><br><span data-ttu-id="c806f-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-620">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-623">
         - PdfFile</span></span><br><span data-ttu-id="c806f-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-624">
         - Selection</span></span><br><span data-ttu-id="c806f-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-625">
         - Settings</span></span><br><span data-ttu-id="c806f-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-626">
         - TableBindings</span></span><br><span data-ttu-id="c806f-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-627">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-628">
         - TextBindings</span></span><br><span data-ttu-id="c806f-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-629">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-631">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c806f-631">Office 2013 on Windows</span></span><br><span data-ttu-id="c806f-632">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-633">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c806f-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c806f-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-636">- BindingEvents</span></span><br><span data-ttu-id="c806f-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-637">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-639">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-640">
         - File</span></span><br><span data-ttu-id="c806f-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-642">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-645">
         - PdfFile</span></span><br><span data-ttu-id="c806f-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-646">
         - Selection</span></span><br><span data-ttu-id="c806f-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-647">
         - Settings</span></span><br><span data-ttu-id="c806f-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-648">
         - TableBindings</span></span><br><span data-ttu-id="c806f-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-649">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-650">
         - TextBindings</span></span><br><span data-ttu-id="c806f-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-651">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-653">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-653">Office on iPad</span></span><br><span data-ttu-id="c806f-654">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-655">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c806f-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c806f-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c806f-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-661">- BindingEvents</span></span><br><span data-ttu-id="c806f-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-662">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-664">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-665">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-665">
         - File</span></span><br><span data-ttu-id="c806f-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-667">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-670">
         - PdfFile</span></span><br><span data-ttu-id="c806f-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-671">
         - Selection</span></span><br><span data-ttu-id="c806f-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-672">
         - Settings</span></span><br><span data-ttu-id="c806f-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-673">
         - TableBindings</span></span><br><span data-ttu-id="c806f-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-674">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-675">
         - TextBindings</span></span><br><span data-ttu-id="c806f-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-676">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-678">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-678">Office on Mac</span></span><br><span data-ttu-id="c806f-679">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-680">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-680">- TaskPane</span></span><br><span data-ttu-id="c806f-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c806f-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c806f-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="c806f-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-688">- BindingEvents</span></span><br><span data-ttu-id="c806f-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-689">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-691">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-692">
         - File</span></span><br><span data-ttu-id="c806f-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-694">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-697">
         - PdfFile</span></span><br><span data-ttu-id="c806f-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-698">
         - Selection</span></span><br><span data-ttu-id="c806f-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-699">
         - Settings</span></span><br><span data-ttu-id="c806f-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-700">
         - TableBindings</span></span><br><span data-ttu-id="c806f-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-701">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-702">
         - TextBindings</span></span><br><span data-ttu-id="c806f-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-703">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-705">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-705">Office 2019 on Mac</span></span><br><span data-ttu-id="c806f-706">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-707">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-707">- TaskPane</span></span><br><span data-ttu-id="c806f-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="c806f-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="c806f-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="c806f-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="c806f-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-714">- BindingEvents</span></span><br><span data-ttu-id="c806f-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-715">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-717">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-718">
         - File</span></span><br><span data-ttu-id="c806f-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-720">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-723">
         - PdfFile</span></span><br><span data-ttu-id="c806f-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-724">
         - Selection</span></span><br><span data-ttu-id="c806f-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-725">
         - Settings</span></span><br><span data-ttu-id="c806f-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-726">
         - TableBindings</span></span><br><span data-ttu-id="c806f-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-727">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-728">
         - TextBindings</span></span><br><span data-ttu-id="c806f-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-729">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-731">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-731">Office 2016 on Mac</span></span><br><span data-ttu-id="c806f-732">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-733">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="c806f-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="c806f-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="c806f-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-737">- BindingEvents</span></span><br><span data-ttu-id="c806f-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-738">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="c806f-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="c806f-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-740">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-741">
         - File</span></span><br><span data-ttu-id="c806f-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-743">
         - MatrixBindings</span></span><br><span data-ttu-id="c806f-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="c806f-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="c806f-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-746">
         - PdfFile</span></span><br><span data-ttu-id="c806f-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-747">
         - Selection</span></span><br><span data-ttu-id="c806f-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-748">
         - Settings</span></span><br><span data-ttu-id="c806f-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-749">
         - TableBindings</span></span><br><span data-ttu-id="c806f-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-750">
         - TableCoercion</span></span><br><span data-ttu-id="c806f-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="c806f-751">
         - TextBindings</span></span><br><span data-ttu-id="c806f-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-752">
         - TextCoercion</span></span><br><span data-ttu-id="c806f-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="c806f-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="c806f-754">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="c806f-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="c806f-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c806f-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c806f-756">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-756">Platform</span></span></th>
    <th><span data-ttu-id="c806f-757">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-757">Extension points</span></span></th>
    <th><span data-ttu-id="c806f-758">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="c806f-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-760">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="c806f-761">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-761">- Content</span></span><br><span data-ttu-id="c806f-762">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-762">
         - TaskPane</span></span><br><span data-ttu-id="c806f-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c806f-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c806f-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-768">- ActiveView</span></span><br><span data-ttu-id="c806f-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-769">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-770">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-771">
         - File</span></span><br><span data-ttu-id="c806f-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-772">
         - PdfFile</span></span><br><span data-ttu-id="c806f-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-773">
         - Selection</span></span><br><span data-ttu-id="c806f-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-774">
         - Settings</span></span><br><span data-ttu-id="c806f-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-776">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="c806f-776">Office on Windows</span></span><br><span data-ttu-id="c806f-777">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-778">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-778">- Content</span></span><br><span data-ttu-id="c806f-779">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-779">
         - TaskPane</span></span><br><span data-ttu-id="c806f-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c806f-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c806f-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-785">- ActiveView</span></span><br><span data-ttu-id="c806f-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-786">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-787">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-788">
         - File</span></span><br><span data-ttu-id="c806f-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-789">
         - PdfFile</span></span><br><span data-ttu-id="c806f-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-790">
         - Selection</span></span><br><span data-ttu-id="c806f-791">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-791">
         - Settings</span></span><br><span data-ttu-id="c806f-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-793">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-793">Office 2019 on Windows</span></span><br><span data-ttu-id="c806f-794">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-795">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-795">- Content</span></span><br><span data-ttu-id="c806f-796">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-796">
         - TaskPane</span></span><br><span data-ttu-id="c806f-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-800">- ActiveView</span></span><br><span data-ttu-id="c806f-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-801">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-802">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-803">
         - File</span></span><br><span data-ttu-id="c806f-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-804">
         - PdfFile</span></span><br><span data-ttu-id="c806f-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-805">
         - Selection</span></span><br><span data-ttu-id="c806f-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-806">
         - Settings</span></span><br><span data-ttu-id="c806f-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-808">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-808">Office 2016 on Windows</span></span><br><span data-ttu-id="c806f-809">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-810">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-810">- Content</span></span><br><span data-ttu-id="c806f-811">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c806f-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c806f-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-814">- ActiveView</span></span><br><span data-ttu-id="c806f-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-815">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-816">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-817">
         - File</span></span><br><span data-ttu-id="c806f-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-818">
         - PdfFile</span></span><br><span data-ttu-id="c806f-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-819">
         - Selection</span></span><br><span data-ttu-id="c806f-820">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-820">
         - Settings</span></span><br><span data-ttu-id="c806f-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-822">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c806f-822">Office 2013 on Windows</span></span><br><span data-ttu-id="c806f-823">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-824">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-824">- Content</span></span><br><span data-ttu-id="c806f-825">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="c806f-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c806f-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c806f-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-828">- ActiveView</span></span><br><span data-ttu-id="c806f-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-829">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-830">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-831">
         - File</span></span><br><span data-ttu-id="c806f-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-832">
         - PdfFile</span></span><br><span data-ttu-id="c806f-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-833">
         - Selection</span></span><br><span data-ttu-id="c806f-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-834">
         - Settings</span></span><br><span data-ttu-id="c806f-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-836">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-836">Office on iPad</span></span><br><span data-ttu-id="c806f-837">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-838">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-838">- Content</span></span><br><span data-ttu-id="c806f-839">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c806f-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-843">- ActiveView</span></span><br><span data-ttu-id="c806f-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-844">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-845">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-846">
         - File</span></span><br><span data-ttu-id="c806f-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-847">
         - PdfFile</span></span><br><span data-ttu-id="c806f-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-848">
         - Selection</span></span><br><span data-ttu-id="c806f-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-849">
         - Settings</span></span><br><span data-ttu-id="c806f-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-851">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="c806f-851">Office on Mac</span></span><br><span data-ttu-id="c806f-852">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="c806f-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="c806f-853">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-853">- Content</span></span><br><span data-ttu-id="c806f-854">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-854">
         - TaskPane</span></span><br><span data-ttu-id="c806f-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="c806f-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="c806f-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="c806f-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="c806f-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-860">- ActiveView</span></span><br><span data-ttu-id="c806f-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-861">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-862">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-863">
         - File</span></span><br><span data-ttu-id="c806f-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-864">
         - PdfFile</span></span><br><span data-ttu-id="c806f-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-865">
         - Selection</span></span><br><span data-ttu-id="c806f-866">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-866">
         - Settings</span></span><br><span data-ttu-id="c806f-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-868">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-868">Office 2019 on Mac</span></span><br><span data-ttu-id="c806f-869">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-870">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-870">- Content</span></span><br><span data-ttu-id="c806f-871">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-871">
         - TaskPane</span></span><br><span data-ttu-id="c806f-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-875">- ActiveView</span></span><br><span data-ttu-id="c806f-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-876">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-877">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-878">
         - File</span></span><br><span data-ttu-id="c806f-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-879">
         - PdfFile</span></span><br><span data-ttu-id="c806f-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-880">
         - Selection</span></span><br><span data-ttu-id="c806f-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-881">
         - Settings</span></span><br><span data-ttu-id="c806f-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-883">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-883">Office 2016 on Mac</span></span><br><span data-ttu-id="c806f-884">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-885">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-885">- Content</span></span><br><span data-ttu-id="c806f-886">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="c806f-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="c806f-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="c806f-889">- ActiveView</span></span><br><span data-ttu-id="c806f-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="c806f-890">
         - CompressedFile</span></span><br><span data-ttu-id="c806f-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-891">
         - DocumentEvents</span></span><br><span data-ttu-id="c806f-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="c806f-892">
         - File</span></span><br><span data-ttu-id="c806f-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="c806f-893">
         - PdfFile</span></span><br><span data-ttu-id="c806f-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-894">
         - Selection</span></span><br><span data-ttu-id="c806f-895">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-895">
         - Settings</span></span><br><span data-ttu-id="c806f-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="c806f-897">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="c806f-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="c806f-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="c806f-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c806f-899">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-899">Platform</span></span></th>
    <th><span data-ttu-id="c806f-900">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-900">Extension points</span></span></th>
    <th><span data-ttu-id="c806f-901">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="c806f-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-903">Office on the web</span><span class="sxs-lookup"><span data-stu-id="c806f-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="c806f-904">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="c806f-904">- Content</span></span><br><span data-ttu-id="c806f-905">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-905">
         - TaskPane</span></span><br><span data-ttu-id="c806f-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="c806f-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="c806f-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="c806f-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="c806f-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="c806f-910">- DocumentEvents</span></span><br><span data-ttu-id="c806f-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="c806f-912">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="c806f-912">
         - Settings</span></span><br><span data-ttu-id="c806f-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="c806f-914">Project</span><span class="sxs-lookup"><span data-stu-id="c806f-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="c806f-915">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="c806f-915">Platform</span></span></th>
    <th><span data-ttu-id="c806f-916">拡張点</span><span class="sxs-lookup"><span data-stu-id="c806f-916">Extension points</span></span></th>
    <th><span data-ttu-id="c806f-917">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="c806f-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="c806f-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-919">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="c806f-919">Office 2019 on Windows</span></span><br><span data-ttu-id="c806f-920">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-921">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-923">- Selection</span></span><br><span data-ttu-id="c806f-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-925">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="c806f-925">Office 2016 on Windows</span></span><br><span data-ttu-id="c806f-926">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-927">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-929">- Selection</span></span><br><span data-ttu-id="c806f-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="c806f-931">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="c806f-931">Office 2013 on Windows</span></span><br><span data-ttu-id="c806f-932">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="c806f-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="c806f-933">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="c806f-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="c806f-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="c806f-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="c806f-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="c806f-935">- Selection</span></span><br><span data-ttu-id="c806f-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="c806f-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="c806f-937">関連項目</span><span class="sxs-lookup"><span data-stu-id="c806f-937">See also</span></span>

- [<span data-ttu-id="c806f-938">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="c806f-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="c806f-939">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="c806f-940">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="c806f-941">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="c806f-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="c806f-942">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="c806f-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="c806f-943">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="c806f-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="c806f-944">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="c806f-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="c806f-945">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="c806f-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="c806f-946">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="c806f-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="c806f-947">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="c806f-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="c806f-948">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="c806f-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="c806f-949">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="c806f-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)