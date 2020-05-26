---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 05/11/2020
localization_priority: Priority
ms.openlocfilehash: 36c6bc6b6348ac988049f9a50127f6dd2f94bf37
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217824"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="1bf1b-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="1bf1b-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="1bf1b-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="1bf1b-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="1bf1b-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="1bf1b-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="1bf1b-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="1bf1b-108">Excel</span><span class="sxs-lookup"><span data-stu-id="1bf1b-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1bf1b-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1bf1b-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1bf1b-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1bf1b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="1bf1b-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-114">- TaskPane</span></span><br><span data-ttu-id="1bf1b-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-115">
        - Content</span></span><br><span data-ttu-id="1bf1b-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1bf1b-116">
        - Custom Functions</span></span><br><span data-ttu-id="1bf1b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="1bf1b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1bf1b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1bf1b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1bf1b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1bf1b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1bf1b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1bf1b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1bf1b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="1bf1b-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="1bf1b-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="1bf1b-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-131">
        - BindingEvents</span></span><br><span data-ttu-id="1bf1b-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-132">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-133">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-134">
        - File</span></span><br><span data-ttu-id="1bf1b-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-135">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-137">
        - Selection</span></span><br><span data-ttu-id="1bf1b-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-138">
        - Settings</span></span><br><span data-ttu-id="1bf1b-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-139">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-140">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-141">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-143">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-143">Office on Windows</span></span><br><span data-ttu-id="1bf1b-144">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-145">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-145">- TaskPane</span></span><br><span data-ttu-id="1bf1b-146">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-146">
        - Content</span></span><br><span data-ttu-id="1bf1b-147">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1bf1b-147">
        - Custom Functions</span></span><br><span data-ttu-id="1bf1b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="1bf1b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="1bf1b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1bf1b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1bf1b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1bf1b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1bf1b-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1bf1b-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1bf1b-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="1bf1b-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="1bf1b-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="1bf1b-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-163">
        - BindingEvents</span></span><br><span data-ttu-id="1bf1b-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-164">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-165">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-166">
        - File</span></span><br><span data-ttu-id="1bf1b-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-167">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-169">
        - Selection</span></span><br><span data-ttu-id="1bf1b-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-170">
        - Settings</span></span><br><span data-ttu-id="1bf1b-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-171">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-172">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-173">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-175">Office 2019 on Windows</span></span><br><span data-ttu-id="1bf1b-176">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1bf1b-177">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-177">- TaskPane</span></span><br><span data-ttu-id="1bf1b-178">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-178">
        - Content</span></span><br><span data-ttu-id="1bf1b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1bf1b-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1bf1b-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1bf1b-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1bf1b-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1bf1b-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1bf1b-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-190">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-191">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-192">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-193">
        - File</span></span><br><span data-ttu-id="1bf1b-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-194">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-196">
        - Selection</span></span><br><span data-ttu-id="1bf1b-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-197">
        - Settings</span></span><br><span data-ttu-id="1bf1b-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-198">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-199">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-200">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-202">Office 2016 on Windows</span></span><br><span data-ttu-id="1bf1b-203">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1bf1b-204">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-204">- TaskPane</span></span><br><span data-ttu-id="1bf1b-205">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-205">
        - Content</span></span></td>
    <td><span data-ttu-id="1bf1b-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1bf1b-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-209">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-210">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-211">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-212">
        - File</span></span><br><span data-ttu-id="1bf1b-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-213">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-215">
        - Selection</span></span><br><span data-ttu-id="1bf1b-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-216">
        - Settings</span></span><br><span data-ttu-id="1bf1b-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-217">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-218">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-219">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1bf1b-221">Office 2013 on Windows</span></span><br><span data-ttu-id="1bf1b-222">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1bf1b-223">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-223">
        - TaskPane</span></span><br><span data-ttu-id="1bf1b-224">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="1bf1b-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1bf1b-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-227">
        - BindingEvents</span></span><br><span data-ttu-id="1bf1b-228">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-228">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-229">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-229">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-230">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-230">
        - File</span></span><br><span data-ttu-id="1bf1b-231">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-231">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-232">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-232">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-233">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-233">
        - Selection</span></span><br><span data-ttu-id="1bf1b-234">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-234">
        - Settings</span></span><br><span data-ttu-id="1bf1b-235">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-235">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-236">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-236">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-237">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-237">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-238">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-238">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-239">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-239">Office on iPad</span></span><br><span data-ttu-id="1bf1b-240">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-240">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1bf1b-241">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-241">- TaskPane</span></span><br><span data-ttu-id="1bf1b-242">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-242">
        - Content</span></span></td>
    <td><span data-ttu-id="1bf1b-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-243">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1bf1b-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1bf1b-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1bf1b-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1bf1b-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1bf1b-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1bf1b-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="1bf1b-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="1bf1b-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-255">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-256">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-256">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-257">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-257">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-258">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-258">
        - File</span></span><br><span data-ttu-id="1bf1b-259">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-259">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-260">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-260">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-261">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-261">
        - Selection</span></span><br><span data-ttu-id="1bf1b-262">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-262">
        - Settings</span></span><br><span data-ttu-id="1bf1b-263">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-263">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-264">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-264">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-265">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-265">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-266">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-266">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-267">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-267">Office on Mac</span></span><br><span data-ttu-id="1bf1b-268">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-268">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1bf1b-269">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-269">- TaskPane</span></span><br><span data-ttu-id="1bf1b-270">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-270">
        - Content</span></span><br><span data-ttu-id="1bf1b-271">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1bf1b-271">
        - Custom Functions</span></span><br><span data-ttu-id="1bf1b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1bf1b-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-273">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1bf1b-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1bf1b-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1bf1b-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1bf1b-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1bf1b-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="1bf1b-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="1bf1b-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="1bf1b-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-286">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="1bf1b-287">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-287">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-288">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-288">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-289">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-289">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-290">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-290">
        - File</span></span><br><span data-ttu-id="1bf1b-291">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-291">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-292">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-292">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-293">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-293">
        - PdfFile</span></span><br><span data-ttu-id="1bf1b-294">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-294">
        - Selection</span></span><br><span data-ttu-id="1bf1b-295">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-295">
        - Settings</span></span><br><span data-ttu-id="1bf1b-296">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-296">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-297">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-297">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-298">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-298">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-299">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-299">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-300">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-300">Office 2019 on Mac</span></span><br><span data-ttu-id="1bf1b-301">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-301">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1bf1b-302">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-302">- TaskPane</span></span><br><span data-ttu-id="1bf1b-303">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-303">
        - Content</span></span><br><span data-ttu-id="1bf1b-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="1bf1b-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-305">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="1bf1b-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="1bf1b-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="1bf1b-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="1bf1b-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="1bf1b-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-314">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-315">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-315">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-316">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-316">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-317">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-317">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-318">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-318">
        - File</span></span><br><span data-ttu-id="1bf1b-319">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-319">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-320">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-320">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-321">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-321">
        - PdfFile</span></span><br><span data-ttu-id="1bf1b-322">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-322">
        - Selection</span></span><br><span data-ttu-id="1bf1b-323">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-323">
        - Settings</span></span><br><span data-ttu-id="1bf1b-324">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-324">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-325">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-325">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-326">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-326">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-327">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-327">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-328">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-328">Office 2016 on Mac</span></span><br><span data-ttu-id="1bf1b-329">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-329">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="1bf1b-330">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-330">- TaskPane</span></span><br><span data-ttu-id="1bf1b-331">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-331">
        - Content</span></span></td>
    <td><span data-ttu-id="1bf1b-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-332">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1bf1b-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-334">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="1bf1b-335">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-335">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-336">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-336">
        - CompressedFile</span></span><br><span data-ttu-id="1bf1b-337">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-337">
        - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-338">
        - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-338">
        - File</span></span><br><span data-ttu-id="1bf1b-339">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-339">
        - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-340">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-340">
        - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-341">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-341">
        - PdfFile</span></span><br><span data-ttu-id="1bf1b-342">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-342">
        - Selection</span></span><br><span data-ttu-id="1bf1b-343">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-343">
        - Settings</span></span><br><span data-ttu-id="1bf1b-344">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-344">
        - TableBindings</span></span><br><span data-ttu-id="1bf1b-345">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-345">
        - TableCoercion</span></span><br><span data-ttu-id="1bf1b-346">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-346">
        - TextBindings</span></span><br><span data-ttu-id="1bf1b-347">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-347">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1bf1b-348">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-348">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="1bf1b-349">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-349">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="1bf1b-350">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-350">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="1bf1b-351">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-351">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="1bf1b-352">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-352">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="1bf1b-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-353"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-354">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-354">Office on the web</span></span></td>
    <td><span data-ttu-id="1bf1b-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1bf1b-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1bf1b-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-357">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-357">Office on Windows</span></span><br><span data-ttu-id="1bf1b-358">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-358">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="1bf1b-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1bf1b-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1bf1b-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-361">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="1bf1b-361">Office for Mac</span></span><br><span data-ttu-id="1bf1b-362">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-362">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="1bf1b-363">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="1bf1b-363">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="1bf1b-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-364">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="1bf1b-365">Outlook</span><span class="sxs-lookup"><span data-stu-id="1bf1b-365">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1bf1b-366">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-366">Platform</span></span></th>
    <th><span data-ttu-id="1bf1b-367">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-367">Extension points</span></span></th>
    <th><span data-ttu-id="1bf1b-368">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-368">API requirement sets</span></span></th>
    <th><span data-ttu-id="1bf1b-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-369"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-370">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-370">Office on the web</span></span><br><span data-ttu-id="1bf1b-371">(モダン)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-371">(modern)</span></span></td>
    <td> <span data-ttu-id="1bf1b-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-372">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-375">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-376">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-377">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1bf1b-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="1bf1b-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="1bf1b-385">利用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-386">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-386">Office on the web</span></span><br><span data-ttu-id="1bf1b-387">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-387">(classic)</span></span></td>
    <td> <span data-ttu-id="1bf1b-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-388">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-391">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-392">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1bf1b-399">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-400">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-400">Office on Windows</span></span><br><span data-ttu-id="1bf1b-401">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-401">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-402">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-405">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1bf1b-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-407">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1bf1b-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="1bf1b-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="1bf1b-416">利用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-417">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-417">Office 2019 on Windows</span></span><br><span data-ttu-id="1bf1b-418">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-419">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-422">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-423">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1bf1b-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-424">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-425">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1bf1b-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="1bf1b-432">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-432">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-433">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-433">Office 2016 on Windows</span></span><br><span data-ttu-id="1bf1b-434">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-434">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-435">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-438">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="1bf1b-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-440">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1bf1b-445">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-446">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1bf1b-446">Office 2013 on Windows</span></span><br><span data-ttu-id="1bf1b-447">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-447">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-448">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="1bf1b-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="1bf1b-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="1bf1b-456">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-457">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-457">Office on iOS</span></span><br><span data-ttu-id="1bf1b-458">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-458">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-459">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-460">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-461">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1bf1b-466">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-466">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-467">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-467">Office on Mac</span></span><br><span data-ttu-id="1bf1b-468">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-468">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-469">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="1bf1b-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="1bf1b-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="1bf1b-482">利用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-482">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-483">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-483">Office 2019 on Mac</span></span><br><span data-ttu-id="1bf1b-484">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-484">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-485">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-488">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-489">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-490">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1bf1b-496">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-496">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-497">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-497">Office 2016 on Mac</span></span><br><span data-ttu-id="1bf1b-498">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-498">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-499">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="1bf1b-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="1bf1b-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-502">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="1bf1b-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-503">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-504">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="1bf1b-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-509">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="1bf1b-510">使用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-510">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-511">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-511">Office on Android</span></span><br><span data-ttu-id="1bf1b-512">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-512">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-513">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="1bf1b-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">予定の開催者 (作成): オンライン会議</a> (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="1bf1b-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-515">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-516">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="1bf1b-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="1bf1b-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="1bf1b-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="1bf1b-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="1bf1b-521">利用不可</span><span class="sxs-lookup"><span data-stu-id="1bf1b-521">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="1bf1b-522">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-522">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1bf1b-523">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="1bf1b-523">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="1bf1b-524">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="1bf1b-524">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="1bf1b-525">Word</span><span class="sxs-lookup"><span data-stu-id="1bf1b-525">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1bf1b-526">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-526">Platform</span></span></th>
    <th><span data-ttu-id="1bf1b-527">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-527">Extension points</span></span></th>
    <th><span data-ttu-id="1bf1b-528">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-528">API requirement sets</span></span></th>
    <th><span data-ttu-id="1bf1b-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-529"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-530">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-530">Office on the web</span></span></td>
    <td> <span data-ttu-id="1bf1b-531">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-531">- TaskPane</span></span><br><span data-ttu-id="1bf1b-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-533">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-538">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-539">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-539">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-540">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-540">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-541">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-541">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-542">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-542">
         - File</span></span><br><span data-ttu-id="1bf1b-543">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-543">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-544">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-544">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-545">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-545">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-546">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-546">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-547">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-547">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-548">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-548">
         - Selection</span></span><br><span data-ttu-id="1bf1b-549">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-549">
         - Settings</span></span><br><span data-ttu-id="1bf1b-550">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-550">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-551">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-551">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-552">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-552">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-553">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-553">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-554">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-554">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-555">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-555">Office on Windows</span></span><br><span data-ttu-id="1bf1b-556">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-556">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-557">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-557">- TaskPane</span></span><br><span data-ttu-id="1bf1b-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-558">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-559">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-562">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-564">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-565">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-565">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-566">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-566">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-567">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-567">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-568">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-568">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-569">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-569">
         - File</span></span><br><span data-ttu-id="1bf1b-570">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-570">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-571">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-571">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-572">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-572">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-573">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-573">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-574">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-574">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-575">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-575">
         - Selection</span></span><br><span data-ttu-id="1bf1b-576">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-576">
         - Settings</span></span><br><span data-ttu-id="1bf1b-577">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-577">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-578">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-578">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-579">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-579">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-580">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-580">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-581">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-581">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-582">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-582">Office 2019 on Windows</span></span><br><span data-ttu-id="1bf1b-583">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-583">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-584">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-584">- TaskPane</span></span><br><span data-ttu-id="1bf1b-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-585">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-586">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-589">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-591">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-592">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-594">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-595">
         - File</span></span><br><span data-ttu-id="1bf1b-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-597">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-600">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-601">
         - Selection</span></span><br><span data-ttu-id="1bf1b-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-602">
         - Settings</span></span><br><span data-ttu-id="1bf1b-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-603">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-604">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-605">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-606">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-608">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-608">Office 2016 on Windows</span></span><br><span data-ttu-id="1bf1b-609">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-610">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-611">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-612">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1bf1b-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-613">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-614">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-614">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-615">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-615">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-616">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-616">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-617">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-617">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-618">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-618">
         - File</span></span><br><span data-ttu-id="1bf1b-619">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-619">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-620">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-620">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-621">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-621">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-622">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-622">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-623">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-623">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-624">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-624">
         - Selection</span></span><br><span data-ttu-id="1bf1b-625">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-625">
         - Settings</span></span><br><span data-ttu-id="1bf1b-626">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-626">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-627">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-627">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-628">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-628">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-629">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-629">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-630">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-630">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-631">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1bf1b-631">Office 2013 on Windows</span></span><br><span data-ttu-id="1bf1b-632">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-632">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-633">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-633">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-634">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1bf1b-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-635">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-636">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-636">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-637">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-637">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-638">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-638">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-639">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-639">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-640">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-640">
         - File</span></span><br><span data-ttu-id="1bf1b-641">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-641">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-642">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-642">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-643">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-643">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-644">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-644">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-645">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-645">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-646">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-646">
         - Selection</span></span><br><span data-ttu-id="1bf1b-647">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-647">
         - Settings</span></span><br><span data-ttu-id="1bf1b-648">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-648">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-649">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-649">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-650">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-650">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-651">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-651">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-652">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-652">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-653">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-653">Office on iPad</span></span><br><span data-ttu-id="1bf1b-654">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-654">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-655">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-655">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-656">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-660">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="1bf1b-661">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-661">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-662">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-662">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-663">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-663">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-664">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-664">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-665">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-665">
         - File</span></span><br><span data-ttu-id="1bf1b-666">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-666">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-667">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-667">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-668">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-668">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-669">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-669">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-670">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-670">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-671">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-671">
         - Selection</span></span><br><span data-ttu-id="1bf1b-672">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-672">
         - Settings</span></span><br><span data-ttu-id="1bf1b-673">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-673">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-674">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-674">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-675">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-675">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-676">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-676">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-677">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-677">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-678">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-678">Office on Mac</span></span><br><span data-ttu-id="1bf1b-679">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-679">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-680">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-680">- TaskPane</span></span><br><span data-ttu-id="1bf1b-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-681">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-682">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="1bf1b-688">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-688">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-689">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-689">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-690">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-690">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-691">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-691">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-692">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-692">
         - File</span></span><br><span data-ttu-id="1bf1b-693">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-693">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-694">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-694">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-695">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-695">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-696">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-696">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-697">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-697">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-698">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-698">
         - Selection</span></span><br><span data-ttu-id="1bf1b-699">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-699">
         - Settings</span></span><br><span data-ttu-id="1bf1b-700">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-700">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-701">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-701">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-702">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-702">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-703">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-703">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-704">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-704">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-705">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-705">Office 2019 on Mac</span></span><br><span data-ttu-id="1bf1b-706">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-706">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-707">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-707">- TaskPane</span></span><br><span data-ttu-id="1bf1b-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-708">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-709">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="1bf1b-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="1bf1b-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-713">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="1bf1b-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-714">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-715">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-717">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-718">
         - File</span></span><br><span data-ttu-id="1bf1b-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-720">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-723">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-724">
         - Selection</span></span><br><span data-ttu-id="1bf1b-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-725">
         - Settings</span></span><br><span data-ttu-id="1bf1b-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-726">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-727">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-728">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-729">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-730">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-731">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-731">Office 2016 on Mac</span></span><br><span data-ttu-id="1bf1b-732">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-732">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-733">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-733">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-734">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-735">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="1bf1b-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-737">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-737">- BindingEvents</span></span><br><span data-ttu-id="1bf1b-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-738">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-739">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="1bf1b-739">
         - CustomXmlParts</span></span><br><span data-ttu-id="1bf1b-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-740">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-741">
         - File</span></span><br><span data-ttu-id="1bf1b-742">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-742">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-743">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-743">
         - MatrixBindings</span></span><br><span data-ttu-id="1bf1b-744">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-744">
         - MatrixCoercion</span></span><br><span data-ttu-id="1bf1b-745">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-745">
         - OoxmlCoercion</span></span><br><span data-ttu-id="1bf1b-746">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-746">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-747">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-747">
         - Selection</span></span><br><span data-ttu-id="1bf1b-748">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-748">
         - Settings</span></span><br><span data-ttu-id="1bf1b-749">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-749">
         - TableBindings</span></span><br><span data-ttu-id="1bf1b-750">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-750">
         - TableCoercion</span></span><br><span data-ttu-id="1bf1b-751">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-751">
         - TextBindings</span></span><br><span data-ttu-id="1bf1b-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-752">
         - TextCoercion</span></span><br><span data-ttu-id="1bf1b-753">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-753">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="1bf1b-754">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-754">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="1bf1b-755">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1bf1b-755">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1bf1b-756">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-756">Platform</span></span></th>
    <th><span data-ttu-id="1bf1b-757">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-757">Extension points</span></span></th>
    <th><span data-ttu-id="1bf1b-758">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-758">API requirement sets</span></span></th>
    <th><span data-ttu-id="1bf1b-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-759"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-760">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-760">Office on the web</span></span></td>
    <td> <span data-ttu-id="1bf1b-761">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-761">- Content</span></span><br><span data-ttu-id="1bf1b-762">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-762">
         - TaskPane</span></span><br><span data-ttu-id="1bf1b-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-763">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-764">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-768">- ActiveView</span></span><br><span data-ttu-id="1bf1b-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-769">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-770">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-771">
         - File</span></span><br><span data-ttu-id="1bf1b-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-772">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-773">
         - Selection</span></span><br><span data-ttu-id="1bf1b-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-774">
         - Settings</span></span><br><span data-ttu-id="1bf1b-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-776">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-776">Office on Windows</span></span><br><span data-ttu-id="1bf1b-777">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-777">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-778">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-778">- Content</span></span><br><span data-ttu-id="1bf1b-779">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-779">
         - TaskPane</span></span><br><span data-ttu-id="1bf1b-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-780">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-781">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-784">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-785">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-785">- ActiveView</span></span><br><span data-ttu-id="1bf1b-786">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-786">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-787">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-787">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-788">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-788">
         - File</span></span><br><span data-ttu-id="1bf1b-789">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-789">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-790">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-790">
         - Selection</span></span><br><span data-ttu-id="1bf1b-791">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-791">
         - Settings</span></span><br><span data-ttu-id="1bf1b-792">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-792">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-793">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-793">Office 2019 on Windows</span></span><br><span data-ttu-id="1bf1b-794">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-794">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-795">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-795">- Content</span></span><br><span data-ttu-id="1bf1b-796">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-796">
         - TaskPane</span></span><br><span data-ttu-id="1bf1b-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-800">- ActiveView</span></span><br><span data-ttu-id="1bf1b-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-801">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-802">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-803">
         - File</span></span><br><span data-ttu-id="1bf1b-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-804">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-805">
         - Selection</span></span><br><span data-ttu-id="1bf1b-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-806">
         - Settings</span></span><br><span data-ttu-id="1bf1b-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-808">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-808">Office 2016 on Windows</span></span><br><span data-ttu-id="1bf1b-809">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-809">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-810">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-810">- Content</span></span><br><span data-ttu-id="1bf1b-811">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-812">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1bf1b-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-814">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-814">- ActiveView</span></span><br><span data-ttu-id="1bf1b-815">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-815">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-816">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-816">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-817">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-817">
         - File</span></span><br><span data-ttu-id="1bf1b-818">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-818">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-819">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-819">
         - Selection</span></span><br><span data-ttu-id="1bf1b-820">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-820">
         - Settings</span></span><br><span data-ttu-id="1bf1b-821">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-821">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-822">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1bf1b-822">Office 2013 on Windows</span></span><br><span data-ttu-id="1bf1b-823">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-823">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-824">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-824">- Content</span></span><br><span data-ttu-id="1bf1b-825">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-825">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="1bf1b-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-826">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1bf1b-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-828">- ActiveView</span></span><br><span data-ttu-id="1bf1b-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-829">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-830">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-831">
         - File</span></span><br><span data-ttu-id="1bf1b-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-832">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-833">
         - Selection</span></span><br><span data-ttu-id="1bf1b-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-834">
         - Settings</span></span><br><span data-ttu-id="1bf1b-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-836">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-836">Office on iPad</span></span><br><span data-ttu-id="1bf1b-837">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-837">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-838">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-838">- Content</span></span><br><span data-ttu-id="1bf1b-839">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-839">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-840">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-843">- ActiveView</span></span><br><span data-ttu-id="1bf1b-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-844">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-845">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-846">
         - File</span></span><br><span data-ttu-id="1bf1b-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-847">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-848">
         - Selection</span></span><br><span data-ttu-id="1bf1b-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-849">
         - Settings</span></span><br><span data-ttu-id="1bf1b-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-851">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="1bf1b-851">Office on Mac</span></span><br><span data-ttu-id="1bf1b-852">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-852">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="1bf1b-853">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-853">- Content</span></span><br><span data-ttu-id="1bf1b-854">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-854">
         - TaskPane</span></span><br><span data-ttu-id="1bf1b-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-855">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-856">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="1bf1b-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-859">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-860">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-860">- ActiveView</span></span><br><span data-ttu-id="1bf1b-861">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-861">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-862">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-862">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-863">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-863">
         - File</span></span><br><span data-ttu-id="1bf1b-864">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-864">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-865">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-865">
         - Selection</span></span><br><span data-ttu-id="1bf1b-866">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-866">
         - Settings</span></span><br><span data-ttu-id="1bf1b-867">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-867">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-868">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-868">Office 2019 on Mac</span></span><br><span data-ttu-id="1bf1b-869">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-869">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-870">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-870">- Content</span></span><br><span data-ttu-id="1bf1b-871">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-871">
         - TaskPane</span></span><br><span data-ttu-id="1bf1b-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-873">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-875">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-875">- ActiveView</span></span><br><span data-ttu-id="1bf1b-876">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-876">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-877">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-877">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-878">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-878">
         - File</span></span><br><span data-ttu-id="1bf1b-879">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-879">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-880">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-880">
         - Selection</span></span><br><span data-ttu-id="1bf1b-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-881">
         - Settings</span></span><br><span data-ttu-id="1bf1b-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-882">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-883">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-883">Office 2016 on Mac</span></span><br><span data-ttu-id="1bf1b-884">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-884">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-885">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-885">- Content</span></span><br><span data-ttu-id="1bf1b-886">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-886">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-887">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="1bf1b-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-888">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-889">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="1bf1b-889">- ActiveView</span></span><br><span data-ttu-id="1bf1b-890">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-890">
         - CompressedFile</span></span><br><span data-ttu-id="1bf1b-891">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-891">
         - DocumentEvents</span></span><br><span data-ttu-id="1bf1b-892">
         - File</span><span class="sxs-lookup"><span data-stu-id="1bf1b-892">
         - File</span></span><br><span data-ttu-id="1bf1b-893">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="1bf1b-893">
         - PdfFile</span></span><br><span data-ttu-id="1bf1b-894">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-894">
         - Selection</span></span><br><span data-ttu-id="1bf1b-895">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-895">
         - Settings</span></span><br><span data-ttu-id="1bf1b-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-896">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="1bf1b-897">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="1bf1b-897">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="1bf1b-898">OneNote</span><span class="sxs-lookup"><span data-stu-id="1bf1b-898">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1bf1b-899">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-899">Platform</span></span></th>
    <th><span data-ttu-id="1bf1b-900">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-900">Extension points</span></span></th>
    <th><span data-ttu-id="1bf1b-901">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-901">API requirement sets</span></span></th>
    <th><span data-ttu-id="1bf1b-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-902"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-903">Office on the web</span><span class="sxs-lookup"><span data-stu-id="1bf1b-903">Office on the web</span></span></td>
    <td> <span data-ttu-id="1bf1b-904">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-904">- Content</span></span><br><span data-ttu-id="1bf1b-905">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-905">
         - TaskPane</span></span><br><span data-ttu-id="1bf1b-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-906">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-907">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="1bf1b-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-909">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-910">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="1bf1b-910">- DocumentEvents</span></span><br><span data-ttu-id="1bf1b-911">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-911">
         - HtmlCoercion</span></span><br><span data-ttu-id="1bf1b-912">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="1bf1b-912">
         - Settings</span></span><br><span data-ttu-id="1bf1b-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="1bf1b-914">Project</span><span class="sxs-lookup"><span data-stu-id="1bf1b-914">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="1bf1b-915">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="1bf1b-915">Platform</span></span></th>
    <th><span data-ttu-id="1bf1b-916">拡張点</span><span class="sxs-lookup"><span data-stu-id="1bf1b-916">Extension points</span></span></th>
    <th><span data-ttu-id="1bf1b-917">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-917">API requirement sets</span></span></th>
    <th><span data-ttu-id="1bf1b-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-918"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-919">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="1bf1b-919">Office 2019 on Windows</span></span><br><span data-ttu-id="1bf1b-920">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-920">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-921">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-921">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-922">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-923">- Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-923">- Selection</span></span><br><span data-ttu-id="1bf1b-924">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-924">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-925">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="1bf1b-925">Office 2016 on Windows</span></span><br><span data-ttu-id="1bf1b-926">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-926">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-927">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-927">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-928">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-929">- Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-929">- Selection</span></span><br><span data-ttu-id="1bf1b-930">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-930">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="1bf1b-931">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="1bf1b-931">Office 2013 on Windows</span></span><br><span data-ttu-id="1bf1b-932">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="1bf1b-932">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="1bf1b-933">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="1bf1b-933">- TaskPane</span></span></td>
    <td> <span data-ttu-id="1bf1b-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="1bf1b-934">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="1bf1b-935">- Selection</span><span class="sxs-lookup"><span data-stu-id="1bf1b-935">- Selection</span></span><br><span data-ttu-id="1bf1b-936">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="1bf1b-936">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="1bf1b-937">関連項目</span><span class="sxs-lookup"><span data-stu-id="1bf1b-937">See also</span></span>

- [<span data-ttu-id="1bf1b-938">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="1bf1b-938">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="1bf1b-939">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-939">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="1bf1b-940">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-940">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="1bf1b-941">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="1bf1b-941">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="1bf1b-942">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="1bf1b-942">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="1bf1b-943">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="1bf1b-943">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="1bf1b-944">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="1bf1b-944">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="1bf1b-945">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="1bf1b-945">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="1bf1b-946">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="1bf1b-946">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="1bf1b-947">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="1bf1b-947">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="1bf1b-948">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="1bf1b-948">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="1bf1b-949">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="1bf1b-949">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)