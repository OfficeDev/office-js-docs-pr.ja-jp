---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 01/23/2020
localization_priority: Priority
ms.openlocfilehash: b30fe872fd89bb02afac99a7838d43d1fbee5464
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554021"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="dd78a-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="dd78a-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="dd78a-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="dd78a-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="dd78a-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="dd78a-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="dd78a-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="dd78a-108">Excel</span><span class="sxs-lookup"><span data-stu-id="dd78a-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dd78a-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dd78a-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dd78a-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dd78a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="dd78a-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-114">- TaskPane</span></span><br><span data-ttu-id="dd78a-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-115">
        - Content</span></span><br><span data-ttu-id="dd78a-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dd78a-116">
        - Custom Functions</span></span><br><span data-ttu-id="dd78a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="dd78a-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dd78a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dd78a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dd78a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dd78a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dd78a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dd78a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dd78a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dd78a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dd78a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="dd78a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="dd78a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-130">
        - BindingEvents</span></span><br><span data-ttu-id="dd78a-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-131">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-132">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-133">
        - File</span></span><br><span data-ttu-id="dd78a-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-134">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-136">
        - Selection</span></span><br><span data-ttu-id="dd78a-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-137">
        - Settings</span></span><br><span data-ttu-id="dd78a-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-138">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-139">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-140">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-142">Office on Windows</span></span><br><span data-ttu-id="dd78a-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-144">- TaskPane</span></span><br><span data-ttu-id="dd78a-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-145">
        - Content</span></span><br><span data-ttu-id="dd78a-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dd78a-146">
        - Custom Functions</span></span><br><span data-ttu-id="dd78a-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="dd78a-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dd78a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dd78a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dd78a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dd78a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dd78a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dd78a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dd78a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dd78a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dd78a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="dd78a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="dd78a-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-161">
        - BindingEvents</span></span><br><span data-ttu-id="dd78a-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-162">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-163">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-164">
        - File</span></span><br><span data-ttu-id="dd78a-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-165">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-167">
        - Selection</span></span><br><span data-ttu-id="dd78a-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-168">
        - Settings</span></span><br><span data-ttu-id="dd78a-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-169">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-170">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-171">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-173">Office 2019 on Windows</span></span><br><span data-ttu-id="dd78a-174">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dd78a-175">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-175">- TaskPane</span></span><br><span data-ttu-id="dd78a-176">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-176">
        - Content</span></span><br><span data-ttu-id="dd78a-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dd78a-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dd78a-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dd78a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dd78a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dd78a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dd78a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dd78a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dd78a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-188">- BindingEvents</span></span><br><span data-ttu-id="dd78a-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-189">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-190">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-191">
        - File</span></span><br><span data-ttu-id="dd78a-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-192">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-194">
        - Selection</span></span><br><span data-ttu-id="dd78a-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-195">
        - Settings</span></span><br><span data-ttu-id="dd78a-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-196">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-197">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-198">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-200">Office 2016 on Windows</span></span><br><span data-ttu-id="dd78a-201">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dd78a-202">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-202">- TaskPane</span></span><br><span data-ttu-id="dd78a-203">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-203">
        - Content</span></span></td>
    <td><span data-ttu-id="dd78a-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dd78a-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-207">- BindingEvents</span></span><br><span data-ttu-id="dd78a-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-208">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-209">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-210">
        - File</span></span><br><span data-ttu-id="dd78a-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-211">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-213">
        - Selection</span></span><br><span data-ttu-id="dd78a-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-214">
        - Settings</span></span><br><span data-ttu-id="dd78a-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-215">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-216">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-217">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dd78a-219">Office 2013 on Windows</span></span><br><span data-ttu-id="dd78a-220">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dd78a-221">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-221">
        - TaskPane</span></span><br><span data-ttu-id="dd78a-222">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="dd78a-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dd78a-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dd78a-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-225">
        - BindingEvents</span></span><br><span data-ttu-id="dd78a-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-226">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-227">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-228">
        - File</span></span><br><span data-ttu-id="dd78a-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-229">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-231">
        - Selection</span></span><br><span data-ttu-id="dd78a-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-232">
        - Settings</span></span><br><span data-ttu-id="dd78a-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-233">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-234">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-235">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-237">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-237">Office on iPad</span></span><br><span data-ttu-id="dd78a-238">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dd78a-239">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-239">- TaskPane</span></span><br><span data-ttu-id="dd78a-240">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-240">
        - Content</span></span></td>
    <td><span data-ttu-id="dd78a-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dd78a-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dd78a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dd78a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dd78a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dd78a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dd78a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dd78a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dd78a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="dd78a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-253">- BindingEvents</span></span><br><span data-ttu-id="dd78a-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-254">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-255">
        - File</span></span><br><span data-ttu-id="dd78a-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-256">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-258">
        - Selection</span></span><br><span data-ttu-id="dd78a-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-259">
        - Settings</span></span><br><span data-ttu-id="dd78a-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-260">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-261">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-262">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-264">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-264">Office on Mac</span></span><br><span data-ttu-id="dd78a-265">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dd78a-266">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-266">- TaskPane</span></span><br><span data-ttu-id="dd78a-267">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-267">
        - Content</span></span><br><span data-ttu-id="dd78a-268">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dd78a-268">
        - Custom Functions</span></span><br><span data-ttu-id="dd78a-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dd78a-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dd78a-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dd78a-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dd78a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dd78a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dd78a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dd78a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dd78a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dd78a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="dd78a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="dd78a-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-283">- BindingEvents</span></span><br><span data-ttu-id="dd78a-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-284">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-285">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-286">
        - File</span></span><br><span data-ttu-id="dd78a-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-287">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-289">
        - PdfFile</span></span><br><span data-ttu-id="dd78a-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-290">
        - Selection</span></span><br><span data-ttu-id="dd78a-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-291">
        - Settings</span></span><br><span data-ttu-id="dd78a-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-292">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-293">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-294">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-296">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-296">Office 2019 on Mac</span></span><br><span data-ttu-id="dd78a-297">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dd78a-298">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-298">- TaskPane</span></span><br><span data-ttu-id="dd78a-299">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-299">
        - Content</span></span><br><span data-ttu-id="dd78a-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dd78a-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dd78a-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dd78a-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dd78a-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dd78a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dd78a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dd78a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dd78a-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-311">- BindingEvents</span></span><br><span data-ttu-id="dd78a-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-312">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-313">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-314">
        - File</span></span><br><span data-ttu-id="dd78a-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-315">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-317">
        - PdfFile</span></span><br><span data-ttu-id="dd78a-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-318">
        - Selection</span></span><br><span data-ttu-id="dd78a-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-319">
        - Settings</span></span><br><span data-ttu-id="dd78a-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-320">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-321">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-322">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-324">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-324">Office 2016 on Mac</span></span><br><span data-ttu-id="dd78a-325">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dd78a-326">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-326">- TaskPane</span></span><br><span data-ttu-id="dd78a-327">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-327">
        - Content</span></span></td>
    <td><span data-ttu-id="dd78a-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dd78a-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dd78a-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dd78a-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-331">- BindingEvents</span></span><br><span data-ttu-id="dd78a-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-332">
        - CompressedFile</span></span><br><span data-ttu-id="dd78a-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-333">
        - DocumentEvents</span></span><br><span data-ttu-id="dd78a-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-334">
        - File</span></span><br><span data-ttu-id="dd78a-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-335">
        - MatrixBindings</span></span><br><span data-ttu-id="dd78a-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-337">
        - PdfFile</span></span><br><span data-ttu-id="dd78a-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-338">
        - Selection</span></span><br><span data-ttu-id="dd78a-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-339">
        - Settings</span></span><br><span data-ttu-id="dd78a-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-340">
        - TableBindings</span></span><br><span data-ttu-id="dd78a-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-341">
        - TableCoercion</span></span><br><span data-ttu-id="dd78a-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-342">
        - TextBindings</span></span><br><span data-ttu-id="dd78a-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dd78a-344">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dd78a-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="dd78a-345">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="dd78a-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dd78a-346">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dd78a-347">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dd78a-348">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dd78a-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-350">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-350">Office on the web</span></span></td>
    <td><span data-ttu-id="dd78a-351">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dd78a-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dd78a-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-353">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-353">Office on Windows</span></span><br><span data-ttu-id="dd78a-354">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dd78a-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dd78a-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dd78a-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="dd78a-357">Office for Mac</span></span><br><span data-ttu-id="dd78a-358">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="dd78a-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="dd78a-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dd78a-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dd78a-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="dd78a-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="dd78a-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dd78a-362">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-362">Platform</span></span></th>
    <th><span data-ttu-id="dd78a-363">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-363">Extension points</span></span></th>
    <th><span data-ttu-id="dd78a-364">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="dd78a-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-366">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-366">Office on the web</span></span><br><span data-ttu-id="dd78a-367">(モダン)</span><span class="sxs-lookup"><span data-stu-id="dd78a-367">(modern)</span></span></td>
    <td> <span data-ttu-id="dd78a-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-368">- Mail Read</span></span><br><span data-ttu-id="dd78a-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-369">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dd78a-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="dd78a-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="dd78a-379">利用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-380">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-380">Office on the web</span></span><br><span data-ttu-id="dd78a-381">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="dd78a-381">(classic)</span></span></td>
    <td> <span data-ttu-id="dd78a-382">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-382">- Mail Read</span></span><br><span data-ttu-id="dd78a-383">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-383">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dd78a-391">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-392">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-392">Office on Windows</span></span><br><span data-ttu-id="dd78a-393">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-394">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-394">- Mail Read</span></span><br><span data-ttu-id="dd78a-395">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-395">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dd78a-397">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dd78a-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dd78a-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dd78a-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="dd78a-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="dd78a-406">利用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-407">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-407">Office 2019 on Windows</span></span><br><span data-ttu-id="dd78a-408">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-409">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-409">- Mail Read</span></span><br><span data-ttu-id="dd78a-410">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-410">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dd78a-412">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dd78a-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dd78a-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dd78a-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dd78a-420">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-421">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-421">Office 2016 on Windows</span></span><br><span data-ttu-id="dd78a-422">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-423">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-423">- Mail Read</span></span><br><span data-ttu-id="dd78a-424">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-424">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dd78a-426">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dd78a-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dd78a-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dd78a-431">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-432">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dd78a-432">Office 2013 on Windows</span></span><br><span data-ttu-id="dd78a-433">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-434">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-434">- Mail Read</span></span><br><span data-ttu-id="dd78a-435">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="dd78a-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="dd78a-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dd78a-440">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-441">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-441">Office on iOS</span></span><br><span data-ttu-id="dd78a-442">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-443">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-443">- Mail Read</span></span><br><span data-ttu-id="dd78a-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dd78a-450">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-451">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-451">Office on Mac</span></span><br><span data-ttu-id="dd78a-452">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-453">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-453">- Mail Read</span></span><br><span data-ttu-id="dd78a-454">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-454">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dd78a-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="dd78a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="dd78a-464">利用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-465">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-465">Office 2019 on Mac</span></span><br><span data-ttu-id="dd78a-466">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-467">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-467">- Mail Read</span></span><br><span data-ttu-id="dd78a-468">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-468">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dd78a-476">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-477">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-477">Office 2016 on Mac</span></span><br><span data-ttu-id="dd78a-478">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-479">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-479">- Mail Read</span></span><br><span data-ttu-id="dd78a-480">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dd78a-480">
      - Mail Compose</span></span><br><span data-ttu-id="dd78a-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dd78a-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dd78a-488">使用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-489">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-489">Office on Android</span></span><br><span data-ttu-id="dd78a-490">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-491">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dd78a-491">- Mail Read</span></span><br><span data-ttu-id="dd78a-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dd78a-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dd78a-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dd78a-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dd78a-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dd78a-498">利用不可</span><span class="sxs-lookup"><span data-stu-id="dd78a-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="dd78a-499">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dd78a-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="dd78a-500">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="dd78a-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="dd78a-501">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dd78a-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="dd78a-502">Word</span><span class="sxs-lookup"><span data-stu-id="dd78a-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dd78a-503">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-503">Platform</span></span></th>
    <th><span data-ttu-id="dd78a-504">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-504">Extension points</span></span></th>
    <th><span data-ttu-id="dd78a-505">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="dd78a-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-507">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="dd78a-508">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-508">- TaskPane</span></span><br><span data-ttu-id="dd78a-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dd78a-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dd78a-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dd78a-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-516">- BindingEvents</span></span><br><span data-ttu-id="dd78a-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-518">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-519">
         - File</span></span><br><span data-ttu-id="dd78a-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-521">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-524">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-525">
         - Selection</span></span><br><span data-ttu-id="dd78a-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-526">
         - Settings</span></span><br><span data-ttu-id="dd78a-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-527">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-528">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-529">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-530">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-532">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-532">Office on Windows</span></span><br><span data-ttu-id="dd78a-533">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-534">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-534">- TaskPane</span></span><br><span data-ttu-id="dd78a-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dd78a-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dd78a-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dd78a-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-542">- BindingEvents</span></span><br><span data-ttu-id="dd78a-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-543">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-545">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-546">
         - File</span></span><br><span data-ttu-id="dd78a-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-548">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-551">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-552">
         - Selection</span></span><br><span data-ttu-id="dd78a-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-553">
         - Settings</span></span><br><span data-ttu-id="dd78a-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-554">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-555">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-556">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-557">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-559">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-559">Office 2019 on Windows</span></span><br><span data-ttu-id="dd78a-560">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-561">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-561">- TaskPane</span></span><br><span data-ttu-id="dd78a-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dd78a-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dd78a-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-568">- BindingEvents</span></span><br><span data-ttu-id="dd78a-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-569">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-571">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-572">
         - File</span></span><br><span data-ttu-id="dd78a-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-574">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-577">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-578">
         - Selection</span></span><br><span data-ttu-id="dd78a-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-579">
         - Settings</span></span><br><span data-ttu-id="dd78a-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-580">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-581">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-582">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-583">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-585">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-585">Office 2016 on Windows</span></span><br><span data-ttu-id="dd78a-586">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-587">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dd78a-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-591">- BindingEvents</span></span><br><span data-ttu-id="dd78a-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-592">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-594">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-595">
         - File</span></span><br><span data-ttu-id="dd78a-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-597">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-600">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-601">
         - Selection</span></span><br><span data-ttu-id="dd78a-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-602">
         - Settings</span></span><br><span data-ttu-id="dd78a-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-603">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-604">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-605">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-606">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-608">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dd78a-608">Office 2013 on Windows</span></span><br><span data-ttu-id="dd78a-609">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-610">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dd78a-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dd78a-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-613">- BindingEvents</span></span><br><span data-ttu-id="dd78a-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-614">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-616">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-617">
         - File</span></span><br><span data-ttu-id="dd78a-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-619">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-622">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-623">
         - Selection</span></span><br><span data-ttu-id="dd78a-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-624">
         - Settings</span></span><br><span data-ttu-id="dd78a-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-625">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-626">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-627">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-628">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-630">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-630">Office on iPad</span></span><br><span data-ttu-id="dd78a-631">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-632">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dd78a-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dd78a-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="dd78a-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-638">- BindingEvents</span></span><br><span data-ttu-id="dd78a-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-639">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-641">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-642">
         - File</span></span><br><span data-ttu-id="dd78a-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-644">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-647">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-648">
         - Selection</span></span><br><span data-ttu-id="dd78a-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-649">
         - Settings</span></span><br><span data-ttu-id="dd78a-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-650">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-651">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-652">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-653">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-655">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-655">Office on Mac</span></span><br><span data-ttu-id="dd78a-656">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-657">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-657">- TaskPane</span></span><br><span data-ttu-id="dd78a-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dd78a-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dd78a-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="dd78a-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-665">- BindingEvents</span></span><br><span data-ttu-id="dd78a-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-666">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-668">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-669">
         - File</span></span><br><span data-ttu-id="dd78a-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-671">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-674">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-675">
         - Selection</span></span><br><span data-ttu-id="dd78a-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-676">
         - Settings</span></span><br><span data-ttu-id="dd78a-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-677">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-678">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-679">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-680">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-682">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-682">Office 2019 on Mac</span></span><br><span data-ttu-id="dd78a-683">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-684">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-684">- TaskPane</span></span><br><span data-ttu-id="dd78a-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dd78a-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dd78a-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="dd78a-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-691">- BindingEvents</span></span><br><span data-ttu-id="dd78a-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-692">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-694">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-695">
         - File</span></span><br><span data-ttu-id="dd78a-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-697">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-700">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-701">
         - Selection</span></span><br><span data-ttu-id="dd78a-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-702">
         - Settings</span></span><br><span data-ttu-id="dd78a-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-703">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-704">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-705">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-706">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-708">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-708">Office 2016 on Mac</span></span><br><span data-ttu-id="dd78a-709">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-710">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dd78a-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dd78a-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dd78a-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-714">- BindingEvents</span></span><br><span data-ttu-id="dd78a-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-715">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dd78a-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="dd78a-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-717">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-718">
         - File</span></span><br><span data-ttu-id="dd78a-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-720">
         - MatrixBindings</span></span><br><span data-ttu-id="dd78a-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="dd78a-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dd78a-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-723">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-724">
         - Selection</span></span><br><span data-ttu-id="dd78a-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-725">
         - Settings</span></span><br><span data-ttu-id="dd78a-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-726">
         - TableBindings</span></span><br><span data-ttu-id="dd78a-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-727">
         - TableCoercion</span></span><br><span data-ttu-id="dd78a-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dd78a-728">
         - TextBindings</span></span><br><span data-ttu-id="dd78a-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-729">
         - TextCoercion</span></span><br><span data-ttu-id="dd78a-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="dd78a-731">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dd78a-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="dd78a-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dd78a-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dd78a-733">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-733">Platform</span></span></th>
    <th><span data-ttu-id="dd78a-734">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-734">Extension points</span></span></th>
    <th><span data-ttu-id="dd78a-735">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="dd78a-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-737">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="dd78a-738">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-738">- Content</span></span><br><span data-ttu-id="dd78a-739">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-739">
         - TaskPane</span></span><br><span data-ttu-id="dd78a-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dd78a-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dd78a-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-745">- ActiveView</span></span><br><span data-ttu-id="dd78a-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-746">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-747">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-748">
         - File</span></span><br><span data-ttu-id="dd78a-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-749">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-750">
         - Selection</span></span><br><span data-ttu-id="dd78a-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-751">
         - Settings</span></span><br><span data-ttu-id="dd78a-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-753">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-753">Office on Windows</span></span><br><span data-ttu-id="dd78a-754">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-755">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-755">- Content</span></span><br><span data-ttu-id="dd78a-756">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-756">
         - TaskPane</span></span><br><span data-ttu-id="dd78a-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dd78a-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dd78a-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-762">- ActiveView</span></span><br><span data-ttu-id="dd78a-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-763">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-764">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-765">
         - File</span></span><br><span data-ttu-id="dd78a-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-766">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-767">
         - Selection</span></span><br><span data-ttu-id="dd78a-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-768">
         - Settings</span></span><br><span data-ttu-id="dd78a-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-770">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-770">Office 2019 on Windows</span></span><br><span data-ttu-id="dd78a-771">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-772">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-772">- Content</span></span><br><span data-ttu-id="dd78a-773">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-773">
         - TaskPane</span></span><br><span data-ttu-id="dd78a-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-777">- ActiveView</span></span><br><span data-ttu-id="dd78a-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-778">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-779">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-780">
         - File</span></span><br><span data-ttu-id="dd78a-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-781">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-782">
         - Selection</span></span><br><span data-ttu-id="dd78a-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-783">
         - Settings</span></span><br><span data-ttu-id="dd78a-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-785">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-785">Office 2016 on Windows</span></span><br><span data-ttu-id="dd78a-786">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-787">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-787">- Content</span></span><br><span data-ttu-id="dd78a-788">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dd78a-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dd78a-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-791">- ActiveView</span></span><br><span data-ttu-id="dd78a-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-792">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-793">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-794">
         - File</span></span><br><span data-ttu-id="dd78a-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-795">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-796">
         - Selection</span></span><br><span data-ttu-id="dd78a-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-797">
         - Settings</span></span><br><span data-ttu-id="dd78a-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-799">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dd78a-799">Office 2013 on Windows</span></span><br><span data-ttu-id="dd78a-800">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-801">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-801">- Content</span></span><br><span data-ttu-id="dd78a-802">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="dd78a-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dd78a-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dd78a-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-805">- ActiveView</span></span><br><span data-ttu-id="dd78a-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-806">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-807">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-808">
         - File</span></span><br><span data-ttu-id="dd78a-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-809">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-810">
         - Selection</span></span><br><span data-ttu-id="dd78a-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-811">
         - Settings</span></span><br><span data-ttu-id="dd78a-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-813">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-813">Office on iPad</span></span><br><span data-ttu-id="dd78a-814">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-815">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-815">- Content</span></span><br><span data-ttu-id="dd78a-816">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dd78a-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-820">- ActiveView</span></span><br><span data-ttu-id="dd78a-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-821">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-822">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-823">
         - File</span></span><br><span data-ttu-id="dd78a-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-824">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-825">
         - Selection</span></span><br><span data-ttu-id="dd78a-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-826">
         - Settings</span></span><br><span data-ttu-id="dd78a-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-828">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dd78a-828">Office on Mac</span></span><br><span data-ttu-id="dd78a-829">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dd78a-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dd78a-830">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-830">- Content</span></span><br><span data-ttu-id="dd78a-831">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-831">
         - TaskPane</span></span><br><span data-ttu-id="dd78a-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dd78a-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dd78a-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dd78a-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-837">- ActiveView</span></span><br><span data-ttu-id="dd78a-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-838">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-839">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-840">
         - File</span></span><br><span data-ttu-id="dd78a-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-841">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-842">
         - Selection</span></span><br><span data-ttu-id="dd78a-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-843">
         - Settings</span></span><br><span data-ttu-id="dd78a-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-845">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-845">Office 2019 on Mac</span></span><br><span data-ttu-id="dd78a-846">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-847">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-847">- Content</span></span><br><span data-ttu-id="dd78a-848">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-848">
         - TaskPane</span></span><br><span data-ttu-id="dd78a-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-852">- ActiveView</span></span><br><span data-ttu-id="dd78a-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-853">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-854">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-855">
         - File</span></span><br><span data-ttu-id="dd78a-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-856">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-857">
         - Selection</span></span><br><span data-ttu-id="dd78a-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-858">
         - Settings</span></span><br><span data-ttu-id="dd78a-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-860">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-860">Office 2016 on Mac</span></span><br><span data-ttu-id="dd78a-861">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-862">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-862">- Content</span></span><br><span data-ttu-id="dd78a-863">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dd78a-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dd78a-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dd78a-866">- ActiveView</span></span><br><span data-ttu-id="dd78a-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-867">
         - CompressedFile</span></span><br><span data-ttu-id="dd78a-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-868">
         - DocumentEvents</span></span><br><span data-ttu-id="dd78a-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="dd78a-869">
         - File</span></span><br><span data-ttu-id="dd78a-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dd78a-870">
         - PdfFile</span></span><br><span data-ttu-id="dd78a-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-871">
         - Selection</span></span><br><span data-ttu-id="dd78a-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-872">
         - Settings</span></span><br><span data-ttu-id="dd78a-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dd78a-874">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dd78a-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="dd78a-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="dd78a-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dd78a-876">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-876">Platform</span></span></th>
    <th><span data-ttu-id="dd78a-877">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-877">Extension points</span></span></th>
    <th><span data-ttu-id="dd78a-878">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="dd78a-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-880">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dd78a-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="dd78a-881">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dd78a-881">- Content</span></span><br><span data-ttu-id="dd78a-882">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-882">
         - TaskPane</span></span><br><span data-ttu-id="dd78a-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dd78a-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="dd78a-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dd78a-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dd78a-887">- DocumentEvents</span></span><br><span data-ttu-id="dd78a-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="dd78a-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dd78a-889">
         - Settings</span></span><br><span data-ttu-id="dd78a-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="dd78a-891">Project</span><span class="sxs-lookup"><span data-stu-id="dd78a-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dd78a-892">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dd78a-892">Platform</span></span></th>
    <th><span data-ttu-id="dd78a-893">拡張点</span><span class="sxs-lookup"><span data-stu-id="dd78a-893">Extension points</span></span></th>
    <th><span data-ttu-id="dd78a-894">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="dd78a-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dd78a-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-896">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dd78a-896">Office 2019 on Windows</span></span><br><span data-ttu-id="dd78a-897">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-898">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-900">- Selection</span></span><br><span data-ttu-id="dd78a-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-902">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dd78a-902">Office 2016 on Windows</span></span><br><span data-ttu-id="dd78a-903">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-904">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-906">- Selection</span></span><br><span data-ttu-id="dd78a-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dd78a-908">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dd78a-908">Office 2013 on Windows</span></span><br><span data-ttu-id="dd78a-909">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dd78a-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dd78a-910">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dd78a-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dd78a-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dd78a-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dd78a-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="dd78a-912">- Selection</span></span><br><span data-ttu-id="dd78a-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dd78a-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="dd78a-914">関連項目</span><span class="sxs-lookup"><span data-stu-id="dd78a-914">See also</span></span>

- [<span data-ttu-id="dd78a-915">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="dd78a-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="dd78a-916">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="dd78a-917">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="dd78a-918">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="dd78a-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="dd78a-919">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="dd78a-919">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="dd78a-920">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="dd78a-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="dd78a-921">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="dd78a-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="dd78a-922">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="dd78a-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="dd78a-923">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="dd78a-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="dd78a-924">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="dd78a-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="dd78a-925">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="dd78a-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="dd78a-926">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="dd78a-926">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)