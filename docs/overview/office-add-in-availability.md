---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 10/30/2019
localization_priority: Priority
ms.openlocfilehash: 3621236ea86410d70d17655450e1f6d32a212823
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901949"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="39ece-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="39ece-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="39ece-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="39ece-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="39ece-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="39ece-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="39ece-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="39ece-108">Excel</span><span class="sxs-lookup"><span data-stu-id="39ece-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="39ece-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="39ece-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="39ece-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="39ece-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="39ece-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-114">- TaskPane</span></span><br><span data-ttu-id="39ece-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-115">
        - Content</span></span><br><span data-ttu-id="39ece-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-116">
        - Custom Functions</span></span><br><span data-ttu-id="39ece-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="39ece-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="39ece-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39ece-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39ece-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39ece-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39ece-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39ece-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39ece-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39ece-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39ece-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39ece-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-128">
        - BindingEvents</span></span><br><span data-ttu-id="39ece-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-129">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-130">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-131">
        - File</span></span><br><span data-ttu-id="39ece-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-132">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-134">
        - Selection</span></span><br><span data-ttu-id="39ece-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-135">
        - Settings</span></span><br><span data-ttu-id="39ece-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-136">
        - TableBindings</span></span><br><span data-ttu-id="39ece-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-137">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-138">
        - TextBindings</span></span><br><span data-ttu-id="39ece-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-140">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="39ece-140">Office on Windows</span></span><br><span data-ttu-id="39ece-141">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-142">- TaskPane</span></span><br><span data-ttu-id="39ece-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-143">
        - Content</span></span><br><span data-ttu-id="39ece-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-144">
        - Custom Functions</span></span><br><span data-ttu-id="39ece-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="39ece-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="39ece-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39ece-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39ece-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39ece-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39ece-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39ece-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39ece-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39ece-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39ece-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39ece-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="39ece-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-158">
        - BindingEvents</span></span><br><span data-ttu-id="39ece-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-159">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-160">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-161">
        - File</span></span><br><span data-ttu-id="39ece-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-162">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-164">
        - Selection</span></span><br><span data-ttu-id="39ece-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-165">
        - Settings</span></span><br><span data-ttu-id="39ece-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-166">
        - TableBindings</span></span><br><span data-ttu-id="39ece-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-167">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-168">
        - TextBindings</span></span><br><span data-ttu-id="39ece-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-170">Office 2019 on Windows</span></span><br><span data-ttu-id="39ece-171">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39ece-172">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-172">- TaskPane</span></span><br><span data-ttu-id="39ece-173">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-173">
        - Content</span></span><br><span data-ttu-id="39ece-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="39ece-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39ece-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39ece-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39ece-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39ece-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39ece-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39ece-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39ece-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-185">- BindingEvents</span></span><br><span data-ttu-id="39ece-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-186">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-187">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-188">
        - File</span></span><br><span data-ttu-id="39ece-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-189">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-191">
        - Selection</span></span><br><span data-ttu-id="39ece-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-192">
        - Settings</span></span><br><span data-ttu-id="39ece-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-193">
        - TableBindings</span></span><br><span data-ttu-id="39ece-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-194">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-195">
        - TextBindings</span></span><br><span data-ttu-id="39ece-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-197">Office 2016 on Windows</span></span><br><span data-ttu-id="39ece-198">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39ece-199">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-199">- TaskPane</span></span><br><span data-ttu-id="39ece-200">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-200">
        - Content</span></span></td>
    <td><span data-ttu-id="39ece-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39ece-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-204">- BindingEvents</span></span><br><span data-ttu-id="39ece-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-205">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-206">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-207">
        - File</span></span><br><span data-ttu-id="39ece-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-208">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-210">
        - Selection</span></span><br><span data-ttu-id="39ece-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-211">
        - Settings</span></span><br><span data-ttu-id="39ece-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-212">
        - TableBindings</span></span><br><span data-ttu-id="39ece-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-213">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-214">
        - TextBindings</span></span><br><span data-ttu-id="39ece-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="39ece-216">Office 2013 on Windows</span></span><br><span data-ttu-id="39ece-217">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39ece-218">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-218">
        - TaskPane</span></span><br><span data-ttu-id="39ece-219">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="39ece-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39ece-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39ece-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-222">
        - BindingEvents</span></span><br><span data-ttu-id="39ece-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-223">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-224">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-225">
        - File</span></span><br><span data-ttu-id="39ece-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-226">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-228">
        - Selection</span></span><br><span data-ttu-id="39ece-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-229">
        - Settings</span></span><br><span data-ttu-id="39ece-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-230">
        - TableBindings</span></span><br><span data-ttu-id="39ece-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-231">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-232">
        - TextBindings</span></span><br><span data-ttu-id="39ece-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-234">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-234">Office on iPad</span></span><br><span data-ttu-id="39ece-235">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="39ece-236">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-236">- TaskPane</span></span><br><span data-ttu-id="39ece-237">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-237">
        - Content</span></span></td>
    <td><span data-ttu-id="39ece-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39ece-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39ece-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39ece-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39ece-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39ece-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39ece-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39ece-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39ece-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39ece-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-249">- BindingEvents</span></span><br><span data-ttu-id="39ece-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-250">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-251">
        - File</span></span><br><span data-ttu-id="39ece-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-252">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-254">
        - Selection</span></span><br><span data-ttu-id="39ece-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-255">
        - Settings</span></span><br><span data-ttu-id="39ece-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-256">
        - TableBindings</span></span><br><span data-ttu-id="39ece-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-257">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-258">
        - TextBindings</span></span><br><span data-ttu-id="39ece-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-260">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-260">Office on Mac</span></span><br><span data-ttu-id="39ece-261">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="39ece-262">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-262">- TaskPane</span></span><br><span data-ttu-id="39ece-263">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-263">
        - Content</span></span><br><span data-ttu-id="39ece-264">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-264">
        - Custom Functions</span></span><br><span data-ttu-id="39ece-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="39ece-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39ece-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39ece-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39ece-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39ece-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39ece-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39ece-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39ece-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="39ece-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="39ece-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="39ece-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-278">- BindingEvents</span></span><br><span data-ttu-id="39ece-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-279">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-280">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-281">
        - File</span></span><br><span data-ttu-id="39ece-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-282">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-284">
        - PdfFile</span></span><br><span data-ttu-id="39ece-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-285">
        - Selection</span></span><br><span data-ttu-id="39ece-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-286">
        - Settings</span></span><br><span data-ttu-id="39ece-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-287">
        - TableBindings</span></span><br><span data-ttu-id="39ece-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-288">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-289">
        - TextBindings</span></span><br><span data-ttu-id="39ece-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-291">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-291">Office 2019 on Mac</span></span><br><span data-ttu-id="39ece-292">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39ece-293">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-293">- TaskPane</span></span><br><span data-ttu-id="39ece-294">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-294">
        - Content</span></span><br><span data-ttu-id="39ece-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="39ece-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="39ece-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="39ece-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="39ece-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="39ece-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="39ece-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="39ece-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="39ece-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-306">- BindingEvents</span></span><br><span data-ttu-id="39ece-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-307">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-308">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-309">
        - File</span></span><br><span data-ttu-id="39ece-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-310">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-312">
        - PdfFile</span></span><br><span data-ttu-id="39ece-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-313">
        - Selection</span></span><br><span data-ttu-id="39ece-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-314">
        - Settings</span></span><br><span data-ttu-id="39ece-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-315">
        - TableBindings</span></span><br><span data-ttu-id="39ece-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-316">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-317">
        - TextBindings</span></span><br><span data-ttu-id="39ece-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-319">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-319">Office 2016 on Mac</span></span><br><span data-ttu-id="39ece-320">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="39ece-321">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-321">- TaskPane</span></span><br><span data-ttu-id="39ece-322">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-322">
        - Content</span></span></td>
    <td><span data-ttu-id="39ece-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="39ece-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39ece-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="39ece-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-326">- BindingEvents</span></span><br><span data-ttu-id="39ece-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-327">
        - CompressedFile</span></span><br><span data-ttu-id="39ece-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-328">
        - DocumentEvents</span></span><br><span data-ttu-id="39ece-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="39ece-329">
        - File</span></span><br><span data-ttu-id="39ece-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-330">
        - MatrixBindings</span></span><br><span data-ttu-id="39ece-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="39ece-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-332">
        - PdfFile</span></span><br><span data-ttu-id="39ece-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-333">
        - Selection</span></span><br><span data-ttu-id="39ece-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-334">
        - Settings</span></span><br><span data-ttu-id="39ece-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-335">
        - TableBindings</span></span><br><span data-ttu-id="39ece-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-336">
        - TableCoercion</span></span><br><span data-ttu-id="39ece-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-337">
        - TextBindings</span></span><br><span data-ttu-id="39ece-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="39ece-339">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="39ece-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="39ece-340">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="39ece-341">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="39ece-342">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="39ece-343">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="39ece-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-345">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-345">Office on the web</span></span></td>
    <td><span data-ttu-id="39ece-346">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="39ece-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-348">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="39ece-348">Office on Windows</span></span><br><span data-ttu-id="39ece-349">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="39ece-350">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="39ece-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-352">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="39ece-352">Office for Mac</span></span><br><span data-ttu-id="39ece-353">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="39ece-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="39ece-354">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="39ece-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="39ece-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="39ece-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="39ece-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39ece-357">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-357">Platform</span></span></th>
    <th><span data-ttu-id="39ece-358">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-358">Extension points</span></span></th>
    <th><span data-ttu-id="39ece-359">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="39ece-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-361">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-361">Office on the web</span></span><br><span data-ttu-id="39ece-362">(モダン)</span><span class="sxs-lookup"><span data-stu-id="39ece-362">(modern)</span></span></td>
    <td> <span data-ttu-id="39ece-363">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-363">- Mail Read</span></span><br><span data-ttu-id="39ece-364">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-364">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39ece-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="39ece-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="39ece-374">利用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-375">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-375">Office on the web</span></span><br><span data-ttu-id="39ece-376">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="39ece-376">(classic)</span></span></td>
    <td> <span data-ttu-id="39ece-377">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-377">- Mail Read</span></span><br><span data-ttu-id="39ece-378">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-378">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="39ece-386">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-387">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="39ece-387">Office on Windows</span></span><br><span data-ttu-id="39ece-388">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-389">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-389">- Mail Read</span></span><br><span data-ttu-id="39ece-390">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-390">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="39ece-392">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="39ece-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="39ece-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39ece-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="39ece-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="39ece-401">利用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-401">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-402">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-402">Office 2019 on Windows</span></span><br><span data-ttu-id="39ece-403">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-403">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-404">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-404">- Mail Read</span></span><br><span data-ttu-id="39ece-405">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-405">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-406">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="39ece-407">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="39ece-407">
      - Modules</span></span></td>
    <td> <span data-ttu-id="39ece-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-408">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39ece-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="39ece-415">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-416">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-416">Office 2016 on Windows</span></span><br><span data-ttu-id="39ece-417">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-418">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-418">- Mail Read</span></span><br><span data-ttu-id="39ece-419">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-419">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-420">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="39ece-421">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="39ece-421">
      - Modules</span></span></td>
    <td> <span data-ttu-id="39ece-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-422">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="39ece-426">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-426">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-427">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="39ece-427">Office 2013 on Windows</span></span><br><span data-ttu-id="39ece-428">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-428">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-429">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-429">- Mail Read</span></span><br><span data-ttu-id="39ece-430">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-430">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="39ece-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-431">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="39ece-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="39ece-435">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-435">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-436">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-436">Office on iOS</span></span><br><span data-ttu-id="39ece-437">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-437">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-438">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-438">- Mail Read</span></span><br><span data-ttu-id="39ece-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-439">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="39ece-445">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-446">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-446">Office on Mac</span></span><br><span data-ttu-id="39ece-447">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-447">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-448">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-448">- Mail Read</span></span><br><span data-ttu-id="39ece-449">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-449">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-450">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="39ece-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="39ece-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="39ece-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="39ece-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="39ece-459">利用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-460">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-460">Office 2019 on Mac</span></span><br><span data-ttu-id="39ece-461">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-462">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-462">- Mail Read</span></span><br><span data-ttu-id="39ece-463">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-463">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="39ece-471">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-472">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-472">Office 2016 on Mac</span></span><br><span data-ttu-id="39ece-473">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-474">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-474">- Mail Read</span></span><br><span data-ttu-id="39ece-475">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="39ece-475">
      - Mail Compose</span></span><br><span data-ttu-id="39ece-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="39ece-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="39ece-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="39ece-483">使用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-484">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-484">Office on Android</span></span><br><span data-ttu-id="39ece-485">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-486">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="39ece-486">- Mail Read</span></span><br><span data-ttu-id="39ece-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="39ece-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="39ece-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="39ece-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="39ece-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="39ece-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="39ece-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="39ece-493">利用不可</span><span class="sxs-lookup"><span data-stu-id="39ece-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="39ece-494">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="39ece-494">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="39ece-495">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="39ece-495">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="39ece-496">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="39ece-496">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="39ece-497">Word</span><span class="sxs-lookup"><span data-stu-id="39ece-497">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39ece-498">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-498">Platform</span></span></th>
    <th><span data-ttu-id="39ece-499">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-499">Extension points</span></span></th>
    <th><span data-ttu-id="39ece-500">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-500">API requirement sets</span></span></th>
    <th><span data-ttu-id="39ece-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-501"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-502">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-502">Office on the web</span></span></td>
    <td> <span data-ttu-id="39ece-503">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-503">- TaskPane</span></span><br><span data-ttu-id="39ece-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-505">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39ece-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39ece-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-510">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39ece-511">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-511">- BindingEvents</span></span><br><span data-ttu-id="39ece-512">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-512">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-513">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-513">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-514">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-514">
         - File</span></span><br><span data-ttu-id="39ece-515">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-515">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-516">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-516">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-517">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-517">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-518">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-518">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-519">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-519">
         - PdfFile</span></span><br><span data-ttu-id="39ece-520">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-520">
         - Selection</span></span><br><span data-ttu-id="39ece-521">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-521">
         - Settings</span></span><br><span data-ttu-id="39ece-522">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-522">
         - TableBindings</span></span><br><span data-ttu-id="39ece-523">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-523">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-524">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-524">
         - TextBindings</span></span><br><span data-ttu-id="39ece-525">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-525">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-526">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-526">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-527">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="39ece-527">Office on Windows</span></span><br><span data-ttu-id="39ece-528">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-528">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-529">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-529">- TaskPane</span></span><br><span data-ttu-id="39ece-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-531">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39ece-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-533">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39ece-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-534">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-535">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-536">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39ece-537">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-537">- BindingEvents</span></span><br><span data-ttu-id="39ece-538">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-538">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-540">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-541">
         - File</span></span><br><span data-ttu-id="39ece-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-543">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-546">
         - PdfFile</span></span><br><span data-ttu-id="39ece-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-547">
         - Selection</span></span><br><span data-ttu-id="39ece-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-548">
         - Settings</span></span><br><span data-ttu-id="39ece-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-549">
         - TableBindings</span></span><br><span data-ttu-id="39ece-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-550">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-551">
         - TextBindings</span></span><br><span data-ttu-id="39ece-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-552">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-553">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-554">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-554">Office 2019 on Windows</span></span><br><span data-ttu-id="39ece-555">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-555">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-556">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-556">- TaskPane</span></span><br><span data-ttu-id="39ece-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39ece-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39ece-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-563">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-563">- BindingEvents</span></span><br><span data-ttu-id="39ece-564">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-564">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-565">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-565">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-566">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-566">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-567">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-567">
         - File</span></span><br><span data-ttu-id="39ece-568">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-568">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-569">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-569">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-570">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-570">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-571">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-571">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-572">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-572">
         - PdfFile</span></span><br><span data-ttu-id="39ece-573">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-573">
         - Selection</span></span><br><span data-ttu-id="39ece-574">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-574">
         - Settings</span></span><br><span data-ttu-id="39ece-575">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-575">
         - TableBindings</span></span><br><span data-ttu-id="39ece-576">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-576">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-577">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-577">
         - TextBindings</span></span><br><span data-ttu-id="39ece-578">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-578">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-579">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-579">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-580">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-580">Office 2016 on Windows</span></span><br><span data-ttu-id="39ece-581">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-581">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-582">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-582">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-583">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39ece-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-586">- BindingEvents</span></span><br><span data-ttu-id="39ece-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-587">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-589">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-590">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-590">
         - File</span></span><br><span data-ttu-id="39ece-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-592">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-595">
         - PdfFile</span></span><br><span data-ttu-id="39ece-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-596">
         - Selection</span></span><br><span data-ttu-id="39ece-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-597">
         - Settings</span></span><br><span data-ttu-id="39ece-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-598">
         - TableBindings</span></span><br><span data-ttu-id="39ece-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-599">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-600">
         - TextBindings</span></span><br><span data-ttu-id="39ece-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-601">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-603">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="39ece-603">Office 2013 on Windows</span></span><br><span data-ttu-id="39ece-604">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-605">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39ece-606">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39ece-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-607">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-608">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-608">- BindingEvents</span></span><br><span data-ttu-id="39ece-609">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-609">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-610">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-610">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-611">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-611">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-612">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-612">
         - File</span></span><br><span data-ttu-id="39ece-613">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-613">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-614">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-614">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-615">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-615">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-616">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-616">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-617">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-617">
         - PdfFile</span></span><br><span data-ttu-id="39ece-618">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-618">
         - Selection</span></span><br><span data-ttu-id="39ece-619">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-619">
         - Settings</span></span><br><span data-ttu-id="39ece-620">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-620">
         - TableBindings</span></span><br><span data-ttu-id="39ece-621">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-621">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-622">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-622">
         - TextBindings</span></span><br><span data-ttu-id="39ece-623">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-623">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-624">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-624">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-625">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-625">Office on iPad</span></span><br><span data-ttu-id="39ece-626">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-626">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-627">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-627">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-628">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39ece-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39ece-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-631">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-632">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="39ece-633">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-633">- BindingEvents</span></span><br><span data-ttu-id="39ece-634">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-634">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-635">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-635">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-636">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-636">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-637">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-637">
         - File</span></span><br><span data-ttu-id="39ece-638">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-638">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-639">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-639">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-640">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-640">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-641">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-641">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-642">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-642">
         - PdfFile</span></span><br><span data-ttu-id="39ece-643">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-643">
         - Selection</span></span><br><span data-ttu-id="39ece-644">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-644">
         - Settings</span></span><br><span data-ttu-id="39ece-645">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-645">
         - TableBindings</span></span><br><span data-ttu-id="39ece-646">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-646">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-647">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-647">
         - TextBindings</span></span><br><span data-ttu-id="39ece-648">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-648">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-649">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-649">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-650">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-650">Office on Mac</span></span><br><span data-ttu-id="39ece-651">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-651">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-652">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-652">- TaskPane</span></span><br><span data-ttu-id="39ece-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-654">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39ece-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39ece-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-658">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-659">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="39ece-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-660">- BindingEvents</span></span><br><span data-ttu-id="39ece-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-661">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-663">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-664">
         - File</span></span><br><span data-ttu-id="39ece-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-666">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-669">
         - PdfFile</span></span><br><span data-ttu-id="39ece-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-670">
         - Selection</span></span><br><span data-ttu-id="39ece-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-671">
         - Settings</span></span><br><span data-ttu-id="39ece-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-672">
         - TableBindings</span></span><br><span data-ttu-id="39ece-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-673">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-674">
         - TextBindings</span></span><br><span data-ttu-id="39ece-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-675">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-677">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-677">Office 2019 on Mac</span></span><br><span data-ttu-id="39ece-678">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-678">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-679">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-679">- TaskPane</span></span><br><span data-ttu-id="39ece-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="39ece-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="39ece-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="39ece-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="39ece-686">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-686">- BindingEvents</span></span><br><span data-ttu-id="39ece-687">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-687">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-688">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-688">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-689">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-689">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-690">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-690">
         - File</span></span><br><span data-ttu-id="39ece-691">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-691">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-692">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-692">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-693">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-693">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-694">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-694">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-695">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-695">
         - PdfFile</span></span><br><span data-ttu-id="39ece-696">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-696">
         - Selection</span></span><br><span data-ttu-id="39ece-697">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-697">
         - Settings</span></span><br><span data-ttu-id="39ece-698">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-698">
         - TableBindings</span></span><br><span data-ttu-id="39ece-699">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-699">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-700">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-700">
         - TextBindings</span></span><br><span data-ttu-id="39ece-701">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-701">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-702">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-702">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-703">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-703">Office 2016 on Mac</span></span><br><span data-ttu-id="39ece-704">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-704">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-705">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-705">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-706">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="39ece-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="39ece-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="39ece-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-708">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-709">- BindingEvents</span></span><br><span data-ttu-id="39ece-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-710">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="39ece-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="39ece-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-712">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-713">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-713">
         - File</span></span><br><span data-ttu-id="39ece-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-715">
         - MatrixBindings</span></span><br><span data-ttu-id="39ece-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="39ece-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="39ece-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-718">
         - PdfFile</span></span><br><span data-ttu-id="39ece-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-719">
         - Selection</span></span><br><span data-ttu-id="39ece-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-720">
         - Settings</span></span><br><span data-ttu-id="39ece-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-721">
         - TableBindings</span></span><br><span data-ttu-id="39ece-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-722">
         - TableCoercion</span></span><br><span data-ttu-id="39ece-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="39ece-723">
         - TextBindings</span></span><br><span data-ttu-id="39ece-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-724">
         - TextCoercion</span></span><br><span data-ttu-id="39ece-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="39ece-725">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="39ece-726">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="39ece-726">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="39ece-727">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="39ece-727">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39ece-728">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-728">Platform</span></span></th>
    <th><span data-ttu-id="39ece-729">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-729">Extension points</span></span></th>
    <th><span data-ttu-id="39ece-730">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-730">API requirement sets</span></span></th>
    <th><span data-ttu-id="39ece-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-731"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-732">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-732">Office on the web</span></span></td>
    <td> <span data-ttu-id="39ece-733">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-733">- Content</span></span><br><span data-ttu-id="39ece-734">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-734">
         - TaskPane</span></span><br><span data-ttu-id="39ece-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-736">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39ece-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-738">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-739">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39ece-740">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-740">- ActiveView</span></span><br><span data-ttu-id="39ece-741">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-741">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-742">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-742">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-743">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-743">
         - File</span></span><br><span data-ttu-id="39ece-744">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-744">
         - PdfFile</span></span><br><span data-ttu-id="39ece-745">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-745">
         - Selection</span></span><br><span data-ttu-id="39ece-746">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-746">
         - Settings</span></span><br><span data-ttu-id="39ece-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-747">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-748">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="39ece-748">Office on Windows</span></span><br><span data-ttu-id="39ece-749">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-749">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-750">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-750">- Content</span></span><br><span data-ttu-id="39ece-751">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-751">
         - TaskPane</span></span><br><span data-ttu-id="39ece-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-753">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39ece-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-755">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-756">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39ece-757">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-757">- ActiveView</span></span><br><span data-ttu-id="39ece-758">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-758">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-759">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-759">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-760">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-760">
         - File</span></span><br><span data-ttu-id="39ece-761">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-761">
         - PdfFile</span></span><br><span data-ttu-id="39ece-762">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-762">
         - Selection</span></span><br><span data-ttu-id="39ece-763">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-763">
         - Settings</span></span><br><span data-ttu-id="39ece-764">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-764">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-765">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-765">Office 2019 on Windows</span></span><br><span data-ttu-id="39ece-766">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-766">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-767">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-767">- Content</span></span><br><span data-ttu-id="39ece-768">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-768">
         - TaskPane</span></span><br><span data-ttu-id="39ece-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-770">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-771">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-772">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-772">- ActiveView</span></span><br><span data-ttu-id="39ece-773">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-773">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-774">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-774">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-775">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-775">
         - File</span></span><br><span data-ttu-id="39ece-776">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-776">
         - PdfFile</span></span><br><span data-ttu-id="39ece-777">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-777">
         - Selection</span></span><br><span data-ttu-id="39ece-778">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-778">
         - Settings</span></span><br><span data-ttu-id="39ece-779">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-779">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-780">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-780">Office 2016 on Windows</span></span><br><span data-ttu-id="39ece-781">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-781">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-782">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-782">- Content</span></span><br><span data-ttu-id="39ece-783">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-783">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39ece-784">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39ece-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-785">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-786">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-786">- ActiveView</span></span><br><span data-ttu-id="39ece-787">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-787">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-788">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-788">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-789">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-789">
         - File</span></span><br><span data-ttu-id="39ece-790">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-790">
         - PdfFile</span></span><br><span data-ttu-id="39ece-791">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-791">
         - Selection</span></span><br><span data-ttu-id="39ece-792">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-792">
         - Settings</span></span><br><span data-ttu-id="39ece-793">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-793">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-794">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="39ece-794">Office 2013 on Windows</span></span><br><span data-ttu-id="39ece-795">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-795">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-796">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-796">- Content</span></span><br><span data-ttu-id="39ece-797">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-797">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="39ece-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39ece-798">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39ece-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-799">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-800">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-800">- ActiveView</span></span><br><span data-ttu-id="39ece-801">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-801">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-802">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-802">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-803">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-803">
         - File</span></span><br><span data-ttu-id="39ece-804">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-804">
         - PdfFile</span></span><br><span data-ttu-id="39ece-805">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-805">
         - Selection</span></span><br><span data-ttu-id="39ece-806">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-806">
         - Settings</span></span><br><span data-ttu-id="39ece-807">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-807">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-808">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-808">Office on iPad</span></span><br><span data-ttu-id="39ece-809">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-809">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-810">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-810">- Content</span></span><br><span data-ttu-id="39ece-811">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-811">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-812">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39ece-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-813">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-814">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-815">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-815">- ActiveView</span></span><br><span data-ttu-id="39ece-816">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-816">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-817">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-817">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-818">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-818">
         - File</span></span><br><span data-ttu-id="39ece-819">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-819">
         - PdfFile</span></span><br><span data-ttu-id="39ece-820">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-820">
         - Selection</span></span><br><span data-ttu-id="39ece-821">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-821">
         - Settings</span></span><br><span data-ttu-id="39ece-822">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-822">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-823">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="39ece-823">Office on Mac</span></span><br><span data-ttu-id="39ece-824">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="39ece-824">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="39ece-825">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-825">- Content</span></span><br><span data-ttu-id="39ece-826">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-826">
         - TaskPane</span></span><br><span data-ttu-id="39ece-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-828">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="39ece-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-830">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="39ece-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="39ece-831">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="39ece-832">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-832">- ActiveView</span></span><br><span data-ttu-id="39ece-833">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-833">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-834">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-834">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-835">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-835">
         - File</span></span><br><span data-ttu-id="39ece-836">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-836">
         - PdfFile</span></span><br><span data-ttu-id="39ece-837">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-837">
         - Selection</span></span><br><span data-ttu-id="39ece-838">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-838">
         - Settings</span></span><br><span data-ttu-id="39ece-839">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-839">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-840">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-840">Office 2019 on Mac</span></span><br><span data-ttu-id="39ece-841">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-841">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-842">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-842">- Content</span></span><br><span data-ttu-id="39ece-843">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-843">
         - TaskPane</span></span><br><span data-ttu-id="39ece-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-845">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-846">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-847">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-847">- ActiveView</span></span><br><span data-ttu-id="39ece-848">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-848">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-849">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-849">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-850">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-850">
         - File</span></span><br><span data-ttu-id="39ece-851">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-851">
         - PdfFile</span></span><br><span data-ttu-id="39ece-852">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-852">
         - Selection</span></span><br><span data-ttu-id="39ece-853">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-853">
         - Settings</span></span><br><span data-ttu-id="39ece-854">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-854">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-855">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-855">Office 2016 on Mac</span></span><br><span data-ttu-id="39ece-856">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-856">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-857">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-857">- Content</span></span><br><span data-ttu-id="39ece-858">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-858">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="39ece-859">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="39ece-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-860">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-861">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="39ece-861">- ActiveView</span></span><br><span data-ttu-id="39ece-862">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="39ece-862">
         - CompressedFile</span></span><br><span data-ttu-id="39ece-863">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-863">
         - DocumentEvents</span></span><br><span data-ttu-id="39ece-864">
         - File</span><span class="sxs-lookup"><span data-stu-id="39ece-864">
         - File</span></span><br><span data-ttu-id="39ece-865">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="39ece-865">
         - PdfFile</span></span><br><span data-ttu-id="39ece-866">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-866">
         - Selection</span></span><br><span data-ttu-id="39ece-867">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-867">
         - Settings</span></span><br><span data-ttu-id="39ece-868">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-868">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="39ece-869">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="39ece-869">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="39ece-870">OneNote</span><span class="sxs-lookup"><span data-stu-id="39ece-870">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39ece-871">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-871">Platform</span></span></th>
    <th><span data-ttu-id="39ece-872">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-872">Extension points</span></span></th>
    <th><span data-ttu-id="39ece-873">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-873">API requirement sets</span></span></th>
    <th><span data-ttu-id="39ece-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-874"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-875">Office on the web</span><span class="sxs-lookup"><span data-stu-id="39ece-875">Office on the web</span></span></td>
    <td> <span data-ttu-id="39ece-876">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="39ece-876">- Content</span></span><br><span data-ttu-id="39ece-877">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-877">
         - TaskPane</span></span><br><span data-ttu-id="39ece-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="39ece-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="39ece-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-879">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="39ece-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-880">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="39ece-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-881">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-882">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="39ece-882">- DocumentEvents</span></span><br><span data-ttu-id="39ece-883">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-883">
         - HtmlCoercion</span></span><br><span data-ttu-id="39ece-884">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="39ece-884">
         - Settings</span></span><br><span data-ttu-id="39ece-885">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-885">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="39ece-886">Project</span><span class="sxs-lookup"><span data-stu-id="39ece-886">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="39ece-887">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="39ece-887">Platform</span></span></th>
    <th><span data-ttu-id="39ece-888">拡張点</span><span class="sxs-lookup"><span data-stu-id="39ece-888">Extension points</span></span></th>
    <th><span data-ttu-id="39ece-889">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-889">API requirement sets</span></span></th>
    <th><span data-ttu-id="39ece-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="39ece-890"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-891">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="39ece-891">Office 2019 on Windows</span></span><br><span data-ttu-id="39ece-892">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-893">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-895">- Selection</span></span><br><span data-ttu-id="39ece-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-897">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="39ece-897">Office 2016 on Windows</span></span><br><span data-ttu-id="39ece-898">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-899">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-901">- Selection</span></span><br><span data-ttu-id="39ece-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-902">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="39ece-903">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="39ece-903">Office 2013 on Windows</span></span><br><span data-ttu-id="39ece-904">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="39ece-904">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="39ece-905">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="39ece-905">- TaskPane</span></span></td>
    <td> <span data-ttu-id="39ece-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="39ece-906">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="39ece-907">- Selection</span><span class="sxs-lookup"><span data-stu-id="39ece-907">- Selection</span></span><br><span data-ttu-id="39ece-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="39ece-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="39ece-909">関連項目</span><span class="sxs-lookup"><span data-stu-id="39ece-909">See also</span></span>

- [<span data-ttu-id="39ece-910">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="39ece-910">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="39ece-911">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-911">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="39ece-912">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-912">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="39ece-913">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="39ece-913">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="39ece-914">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="39ece-914">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="39ece-915">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="39ece-915">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="39ece-916">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="39ece-916">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="39ece-917">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="39ece-917">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="39ece-918">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="39ece-918">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="39ece-919">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="39ece-919">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="39ece-920">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="39ece-920">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
