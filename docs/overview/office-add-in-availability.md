---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: 1e368fe21a1bcdb2a7f44c88ce8e881605fa96f2
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395653"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="dfbc9-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="dfbc9-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="dfbc9-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="dfbc9-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="dfbc9-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="dfbc9-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="dfbc9-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="dfbc9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="dfbc9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dfbc9-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dfbc9-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dfbc9-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dfbc9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="dfbc9-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-114">- TaskPane</span></span><br><span data-ttu-id="dfbc9-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-115">
        - Content</span></span><br><span data-ttu-id="dfbc9-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-116">
        - Custom Functions</span></span><br><span data-ttu-id="dfbc9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="dfbc9-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dfbc9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dfbc9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dfbc9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dfbc9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dfbc9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dfbc9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dfbc9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-128">
        - BindingEvents</span></span><br><span data-ttu-id="dfbc9-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-129">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-130">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-131">
        - File</span></span><br><span data-ttu-id="dfbc9-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-132">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-134">
        - Selection</span></span><br><span data-ttu-id="dfbc9-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-135">
        - Settings</span></span><br><span data-ttu-id="dfbc9-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-136">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-137">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-138">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-140">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-140">Office on Windows</span></span><br><span data-ttu-id="dfbc9-141">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-142">- TaskPane</span></span><br><span data-ttu-id="dfbc9-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-143">
        - Content</span></span><br><span data-ttu-id="dfbc9-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-144">
        - Custom Functions</span></span><br><span data-ttu-id="dfbc9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="dfbc9-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="dfbc9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dfbc9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dfbc9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dfbc9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dfbc9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dfbc9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dfbc9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="dfbc9-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-158">
        - BindingEvents</span></span><br><span data-ttu-id="dfbc9-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-159">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-160">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-161">
        - File</span></span><br><span data-ttu-id="dfbc9-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-162">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-164">
        - Selection</span></span><br><span data-ttu-id="dfbc9-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-165">
        - Settings</span></span><br><span data-ttu-id="dfbc9-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-166">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-167">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-168">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-170">Office 2019 on Windows</span></span><br><span data-ttu-id="dfbc9-171">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dfbc9-172">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-172">- TaskPane</span></span><br><span data-ttu-id="dfbc9-173">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-173">
        - Content</span></span><br><span data-ttu-id="dfbc9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dfbc9-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dfbc9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dfbc9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dfbc9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dfbc9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dfbc9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-185">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-186">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-187">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-188">
        - File</span></span><br><span data-ttu-id="dfbc9-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-189">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-191">
        - Selection</span></span><br><span data-ttu-id="dfbc9-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-192">
        - Settings</span></span><br><span data-ttu-id="dfbc9-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-193">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-194">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-195">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-197">Office 2016 on Windows</span></span><br><span data-ttu-id="dfbc9-198">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dfbc9-199">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-199">- TaskPane</span></span><br><span data-ttu-id="dfbc9-200">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-200">
        - Content</span></span></td>
    <td><span data-ttu-id="dfbc9-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dfbc9-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-204">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-205">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-206">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-207">
        - File</span></span><br><span data-ttu-id="dfbc9-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-208">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-210">
        - Selection</span></span><br><span data-ttu-id="dfbc9-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-211">
        - Settings</span></span><br><span data-ttu-id="dfbc9-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-212">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-213">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-214">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfbc9-216">Office 2013 on Windows</span></span><br><span data-ttu-id="dfbc9-217">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dfbc9-218">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-218">
        - TaskPane</span></span><br><span data-ttu-id="dfbc9-219">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="dfbc9-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dfbc9-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-222">
        - BindingEvents</span></span><br><span data-ttu-id="dfbc9-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-223">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-224">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-225">
        - File</span></span><br><span data-ttu-id="dfbc9-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-226">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-228">
        - Selection</span></span><br><span data-ttu-id="dfbc9-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-229">
        - Settings</span></span><br><span data-ttu-id="dfbc9-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-230">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-231">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-232">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-234">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-234">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="dfbc9-235">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dfbc9-236">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-236">- TaskPane</span></span><br><span data-ttu-id="dfbc9-237">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-237">
        - Content</span></span><br><span data-ttu-id="dfbc9-238">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-238">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dfbc9-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-239">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dfbc9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dfbc9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dfbc9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dfbc9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dfbc9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dfbc9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-250">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-250">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-251">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-251">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-252">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-252">
        - File</span></span><br><span data-ttu-id="dfbc9-253">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-253">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-254">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-254">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-255">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-255">
        - Selection</span></span><br><span data-ttu-id="dfbc9-256">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-256">
        - Settings</span></span><br><span data-ttu-id="dfbc9-257">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-257">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-258">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-258">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-259">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-259">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-260">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-260">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-261">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-261">Office apps on Mac</span></span><br><span data-ttu-id="dfbc9-262">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-262">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dfbc9-263">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-263">- TaskPane</span></span><br><span data-ttu-id="dfbc9-264">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-264">
        - Content</span></span><br><span data-ttu-id="dfbc9-265">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-265">
        - Custom Functions</span></span><br><span data-ttu-id="dfbc9-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-266">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dfbc9-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-267">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dfbc9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dfbc9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dfbc9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dfbc9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dfbc9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="dfbc9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="dfbc9-279">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-279">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-280">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-280">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-281">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-281">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-282">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-282">
        - File</span></span><br><span data-ttu-id="dfbc9-283">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-283">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-284">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-284">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-285">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-285">
        - PdfFile</span></span><br><span data-ttu-id="dfbc9-286">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-286">
        - Selection</span></span><br><span data-ttu-id="dfbc9-287">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-287">
        - Settings</span></span><br><span data-ttu-id="dfbc9-288">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-288">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-289">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-289">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-290">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-290">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-291">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-291">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-292">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-292">Office 2019 for Mac</span></span><br><span data-ttu-id="dfbc9-293">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-293">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dfbc9-294">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-294">- TaskPane</span></span><br><span data-ttu-id="dfbc9-295">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-295">
        - Content</span></span><br><span data-ttu-id="dfbc9-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-296">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="dfbc9-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-297">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="dfbc9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="dfbc9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="dfbc9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="dfbc9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="dfbc9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-307">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-307">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-308">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-308">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-309">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-309">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-310">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-310">
        - File</span></span><br><span data-ttu-id="dfbc9-311">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-311">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-312">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-312">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-313">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-313">
        - PdfFile</span></span><br><span data-ttu-id="dfbc9-314">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-314">
        - Selection</span></span><br><span data-ttu-id="dfbc9-315">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-315">
        - Settings</span></span><br><span data-ttu-id="dfbc9-316">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-316">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-317">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-317">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-318">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-318">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-319">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-319">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-320">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-320">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dfbc9-321">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-321">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="dfbc9-322">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-322">- TaskPane</span></span><br><span data-ttu-id="dfbc9-323">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-323">
        - Content</span></span></td>
    <td><span data-ttu-id="dfbc9-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-324">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dfbc9-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-326">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="dfbc9-327">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-327">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-328">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-328">
        - CompressedFile</span></span><br><span data-ttu-id="dfbc9-329">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-329">
        - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-330">
        - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-330">
        - File</span></span><br><span data-ttu-id="dfbc9-331">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-331">
        - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-332">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-332">
        - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-333">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-333">
        - PdfFile</span></span><br><span data-ttu-id="dfbc9-334">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-334">
        - Selection</span></span><br><span data-ttu-id="dfbc9-335">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-335">
        - Settings</span></span><br><span data-ttu-id="dfbc9-336">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-336">
        - TableBindings</span></span><br><span data-ttu-id="dfbc9-337">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-337">
        - TableCoercion</span></span><br><span data-ttu-id="dfbc9-338">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-338">
        - TextBindings</span></span><br><span data-ttu-id="dfbc9-339">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-339">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dfbc9-340">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-340">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="dfbc9-341">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-341">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="dfbc9-342">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-342">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="dfbc9-343">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-343">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="dfbc9-344">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-344">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="dfbc9-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-345"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-346">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-346">Office on the web</span></span></td>
    <td><span data-ttu-id="dfbc9-347">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-347">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dfbc9-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-348">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-349">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-349">Office on Windows</span></span><br><span data-ttu-id="dfbc9-350">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-350">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="dfbc9-351">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dfbc9-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-353">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="dfbc9-353">Office for Mac</span></span><br><span data-ttu-id="dfbc9-354">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-354">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="dfbc9-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="dfbc9-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="dfbc9-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="dfbc9-357">Outlook</span><span class="sxs-lookup"><span data-stu-id="dfbc9-357">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dfbc9-358">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-358">Platform</span></span></th>
    <th><span data-ttu-id="dfbc9-359">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-359">Extension points</span></span></th>
    <th><span data-ttu-id="dfbc9-360">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-360">API requirement sets</span></span></th>
    <th><span data-ttu-id="dfbc9-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-361"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-362">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-362">Office on the web</span></span><br><span data-ttu-id="dfbc9-363">(モダン)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-363">Modern</span></span></td>
    <td> <span data-ttu-id="dfbc9-364">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-364">- Mail Read</span></span><br><span data-ttu-id="dfbc9-365">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-365">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-366">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-367">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dfbc9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dfbc9-374">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-374">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-375">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-375">Office on the web</span></span><br><span data-ttu-id="dfbc9-376">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-376">Classic</span></span></td>
    <td> <span data-ttu-id="dfbc9-377">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-377">- Mail Read</span></span><br><span data-ttu-id="dfbc9-378">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-378">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-379">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-380">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dfbc9-386">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-386">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-387">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-387">Office on Windows</span></span><br><span data-ttu-id="dfbc9-388">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-388">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-389">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-389">- Mail Read</span></span><br><span data-ttu-id="dfbc9-390">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-390">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dfbc9-392">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dfbc9-392">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dfbc9-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-393">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dfbc9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dfbc9-400">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-400">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-401">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-401">Office 2019 on Windows</span></span><br><span data-ttu-id="dfbc9-402">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-402">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-403">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-403">- Mail Read</span></span><br><span data-ttu-id="dfbc9-404">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-404">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dfbc9-406">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dfbc9-406">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dfbc9-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dfbc9-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dfbc9-414">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-415">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-415">Office 2016 on Windows</span></span><br><span data-ttu-id="dfbc9-416">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-416">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-417">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-417">- Mail Read</span></span><br><span data-ttu-id="dfbc9-418">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-418">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="dfbc9-420">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="dfbc9-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="dfbc9-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dfbc9-425">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-426">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfbc9-426">Office 2013 on Windows</span></span><br><span data-ttu-id="dfbc9-427">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-427">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-428">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-428">- Mail Read</span></span><br><span data-ttu-id="dfbc9-429">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-429">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="dfbc9-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-430">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="dfbc9-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="dfbc9-434">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-434">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-435">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-435">Office apps on iOS</span></span><br><span data-ttu-id="dfbc9-436">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-436">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-437">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-437">- Mail Read</span></span><br><span data-ttu-id="dfbc9-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-439">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dfbc9-444">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-445">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-445">Office apps on Mac</span></span><br><span data-ttu-id="dfbc9-446">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-446">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-447">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-447">- Mail Read</span></span><br><span data-ttu-id="dfbc9-448">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-448">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-449">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-450">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="dfbc9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="dfbc9-457">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-457">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-458">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-458">Office 2019 for Mac</span></span><br><span data-ttu-id="dfbc9-459">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-459">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-460">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-460">- Mail Read</span></span><br><span data-ttu-id="dfbc9-461">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-461">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-462">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-463">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dfbc9-469">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-470">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-470">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dfbc9-471">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-471">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-472">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-472">- Mail Read</span></span><br><span data-ttu-id="dfbc9-473">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="dfbc9-473">
      - Mail Compose</span></span><br><span data-ttu-id="dfbc9-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-474">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-475">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="dfbc9-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="dfbc9-481">使用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-482">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-482">Office apps on Android</span></span><br><span data-ttu-id="dfbc9-483">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-483">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-484">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="dfbc9-484">- Mail Read</span></span><br><span data-ttu-id="dfbc9-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="dfbc9-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="dfbc9-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="dfbc9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="dfbc9-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="dfbc9-491">利用不可</span><span class="sxs-lookup"><span data-stu-id="dfbc9-491">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="dfbc9-492">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-492">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="dfbc9-493">Word</span><span class="sxs-lookup"><span data-stu-id="dfbc9-493">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dfbc9-494">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-494">Platform</span></span></th>
    <th><span data-ttu-id="dfbc9-495">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-495">Extension points</span></span></th>
    <th><span data-ttu-id="dfbc9-496">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-496">API requirement sets</span></span></th>
    <th><span data-ttu-id="dfbc9-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-497"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-498">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-498">Office on the web</span></span></td>
    <td> <span data-ttu-id="dfbc9-499">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-499">- TaskPane</span></span><br><span data-ttu-id="dfbc9-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-500">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-501">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-507">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-507">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-508">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-508">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-509">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-509">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-510">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-510">
         - File</span></span><br><span data-ttu-id="dfbc9-511">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-511">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-512">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-512">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-513">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-513">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-514">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-514">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-515">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-515">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-516">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-516">
         - Selection</span></span><br><span data-ttu-id="dfbc9-517">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-517">
         - Settings</span></span><br><span data-ttu-id="dfbc9-518">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-518">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-519">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-519">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-520">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-520">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-521">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-521">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-522">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-522">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-523">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-523">Office on Windows</span></span><br><span data-ttu-id="dfbc9-524">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-524">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-525">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-525">- TaskPane</span></span><br><span data-ttu-id="dfbc9-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-526">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-527">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-528">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-531">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-533">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-533">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-534">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-534">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-536">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-537">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-537">
         - File</span></span><br><span data-ttu-id="dfbc9-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-539">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-542">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-543">
         - Selection</span></span><br><span data-ttu-id="dfbc9-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-544">
         - Settings</span></span><br><span data-ttu-id="dfbc9-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-545">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-546">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-547">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-548">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-549">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-550">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-550">Office 2019 on Windows</span></span><br><span data-ttu-id="dfbc9-551">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-551">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-552">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-552">- TaskPane</span></span><br><span data-ttu-id="dfbc9-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-559">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-559">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-560">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-560">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-561">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-561">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-562">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-562">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-563">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-563">
         - File</span></span><br><span data-ttu-id="dfbc9-564">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-564">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-565">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-565">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-566">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-566">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-567">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-567">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-568">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-568">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-569">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-569">
         - Selection</span></span><br><span data-ttu-id="dfbc9-570">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-570">
         - Settings</span></span><br><span data-ttu-id="dfbc9-571">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-571">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-572">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-572">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-573">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-573">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-574">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-574">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-575">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-575">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-576">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-576">Office 2016 on Windows</span></span><br><span data-ttu-id="dfbc9-577">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-577">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-578">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-578">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-579">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dfbc9-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-581">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-582">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-582">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-583">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-583">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-584">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-584">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-585">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-585">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-586">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-586">
         - File</span></span><br><span data-ttu-id="dfbc9-587">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-587">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-588">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-588">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-589">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-589">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-590">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-590">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-591">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-591">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-592">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-592">
         - Selection</span></span><br><span data-ttu-id="dfbc9-593">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-593">
         - Settings</span></span><br><span data-ttu-id="dfbc9-594">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-594">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-595">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-595">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-596">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-596">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-597">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-597">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-598">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-598">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-599">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfbc9-599">Office 2013 on Windows</span></span><br><span data-ttu-id="dfbc9-600">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-600">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-601">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-601">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-602">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dfbc9-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-603">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-604">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-604">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-605">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-605">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-606">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-606">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-607">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-607">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-608">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-608">
         - File</span></span><br><span data-ttu-id="dfbc9-609">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-609">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-610">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-610">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-611">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-611">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-612">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-612">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-613">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-613">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-614">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-614">
         - Selection</span></span><br><span data-ttu-id="dfbc9-615">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-615">
         - Settings</span></span><br><span data-ttu-id="dfbc9-616">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-616">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-617">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-617">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-618">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-618">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-619">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-619">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-620">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-620">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-621">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-621">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="dfbc9-622">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-622">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-623">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-623">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-624">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-625">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="dfbc9-629">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-629">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-630">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-630">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-631">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-631">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-632">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-632">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-633">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-633">
         - File</span></span><br><span data-ttu-id="dfbc9-634">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-634">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-635">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-635">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-636">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-636">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-637">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-637">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-638">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-638">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-639">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-639">
         - Selection</span></span><br><span data-ttu-id="dfbc9-640">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-640">
         - Settings</span></span><br><span data-ttu-id="dfbc9-641">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-641">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-642">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-642">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-643">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-643">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-644">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-644">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-645">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-645">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-646">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-646">Office apps on Mac</span></span><br><span data-ttu-id="dfbc9-647">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-647">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-648">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-648">- TaskPane</span></span><br><span data-ttu-id="dfbc9-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-649">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-650">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-651">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="dfbc9-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-656">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-657">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-659">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-660">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-660">
         - File</span></span><br><span data-ttu-id="dfbc9-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-662">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-665">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-666">
         - Selection</span></span><br><span data-ttu-id="dfbc9-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-667">
         - Settings</span></span><br><span data-ttu-id="dfbc9-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-668">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-669">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-670">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-671">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-673">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-673">Office 2019 for Mac</span></span><br><span data-ttu-id="dfbc9-674">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-674">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-675">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-675">- TaskPane</span></span><br><span data-ttu-id="dfbc9-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="dfbc9-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="dfbc9-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="dfbc9-682">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-682">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-683">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-683">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-684">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-684">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-685">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-685">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-686">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-686">
         - File</span></span><br><span data-ttu-id="dfbc9-687">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-687">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-688">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-688">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-689">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-689">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-690">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-690">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-691">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-691">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-692">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-692">
         - Selection</span></span><br><span data-ttu-id="dfbc9-693">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-693">
         - Settings</span></span><br><span data-ttu-id="dfbc9-694">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-694">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-695">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-695">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-696">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-696">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-697">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-697">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-698">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-698">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-699">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-699">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dfbc9-700">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-700">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-701">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-701">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-702">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-703">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="dfbc9-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-704">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-705">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-705">- BindingEvents</span></span><br><span data-ttu-id="dfbc9-706">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-706">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-707">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="dfbc9-707">
         - CustomXmlParts</span></span><br><span data-ttu-id="dfbc9-708">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-708">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-709">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-709">
         - File</span></span><br><span data-ttu-id="dfbc9-710">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-710">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-711">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-711">
         - MatrixBindings</span></span><br><span data-ttu-id="dfbc9-712">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-712">
         - MatrixCoercion</span></span><br><span data-ttu-id="dfbc9-713">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-713">
         - OoxmlCoercion</span></span><br><span data-ttu-id="dfbc9-714">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-714">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-715">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-715">
         - Selection</span></span><br><span data-ttu-id="dfbc9-716">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-716">
         - Settings</span></span><br><span data-ttu-id="dfbc9-717">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-717">
         - TableBindings</span></span><br><span data-ttu-id="dfbc9-718">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-718">
         - TableCoercion</span></span><br><span data-ttu-id="dfbc9-719">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-719">
         - TextBindings</span></span><br><span data-ttu-id="dfbc9-720">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-720">
         - TextCoercion</span></span><br><span data-ttu-id="dfbc9-721">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-721">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="dfbc9-722">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-722">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="dfbc9-723">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="dfbc9-723">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dfbc9-724">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-724">Platform</span></span></th>
    <th><span data-ttu-id="dfbc9-725">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-725">Extension points</span></span></th>
    <th><span data-ttu-id="dfbc9-726">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-726">API requirement sets</span></span></th>
    <th><span data-ttu-id="dfbc9-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-727"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-728">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-728">Office on the web</span></span></td>
    <td> <span data-ttu-id="dfbc9-729">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-729">- Content</span></span><br><span data-ttu-id="dfbc9-730">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-730">
         - TaskPane</span></span><br><span data-ttu-id="dfbc9-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-731">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-732">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-736">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-736">- ActiveView</span></span><br><span data-ttu-id="dfbc9-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-737">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-738">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-738">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-739">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-739">
         - File</span></span><br><span data-ttu-id="dfbc9-740">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-740">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-741">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-741">
         - Selection</span></span><br><span data-ttu-id="dfbc9-742">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-742">
         - Settings</span></span><br><span data-ttu-id="dfbc9-743">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-743">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-744">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-744">Office on Windows</span></span><br><span data-ttu-id="dfbc9-745">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-745">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-746">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-746">- Content</span></span><br><span data-ttu-id="dfbc9-747">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-747">
         - TaskPane</span></span><br><span data-ttu-id="dfbc9-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-748">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-749">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-753">- ActiveView</span></span><br><span data-ttu-id="dfbc9-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-754">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-755">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-756">
         - File</span></span><br><span data-ttu-id="dfbc9-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-757">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-758">
         - Selection</span></span><br><span data-ttu-id="dfbc9-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-759">
         - Settings</span></span><br><span data-ttu-id="dfbc9-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-761">Office 2019 on Windows</span></span><br><span data-ttu-id="dfbc9-762">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-763">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-763">- Content</span></span><br><span data-ttu-id="dfbc9-764">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-764">
         - TaskPane</span></span><br><span data-ttu-id="dfbc9-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-768">- ActiveView</span></span><br><span data-ttu-id="dfbc9-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-769">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-770">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-771">
         - File</span></span><br><span data-ttu-id="dfbc9-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-772">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-773">
         - Selection</span></span><br><span data-ttu-id="dfbc9-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-774">
         - Settings</span></span><br><span data-ttu-id="dfbc9-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-776">Office 2016 on Windows</span></span><br><span data-ttu-id="dfbc9-777">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-778">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-778">- Content</span></span><br><span data-ttu-id="dfbc9-779">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dfbc9-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-782">- ActiveView</span></span><br><span data-ttu-id="dfbc9-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-783">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-784">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-785">
         - File</span></span><br><span data-ttu-id="dfbc9-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-786">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-787">
         - Selection</span></span><br><span data-ttu-id="dfbc9-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-788">
         - Settings</span></span><br><span data-ttu-id="dfbc9-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfbc9-790">Office 2013 on Windows</span></span><br><span data-ttu-id="dfbc9-791">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-792">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-792">- Content</span></span><br><span data-ttu-id="dfbc9-793">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="dfbc9-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dfbc9-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-796">- ActiveView</span></span><br><span data-ttu-id="dfbc9-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-797">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-798">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-799">
         - File</span></span><br><span data-ttu-id="dfbc9-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-800">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-801">
         - Selection</span></span><br><span data-ttu-id="dfbc9-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-802">
         - Settings</span></span><br><span data-ttu-id="dfbc9-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-804">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="dfbc9-805">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-806">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-806">- Content</span></span><br><span data-ttu-id="dfbc9-807">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-808">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-811">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-811">- ActiveView</span></span><br><span data-ttu-id="dfbc9-812">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-812">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-813">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-813">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-814">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-814">
         - File</span></span><br><span data-ttu-id="dfbc9-815">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-815">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-816">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-816">
         - Selection</span></span><br><span data-ttu-id="dfbc9-817">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-817">
         - Settings</span></span><br><span data-ttu-id="dfbc9-818">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-818">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-819">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="dfbc9-819">Office apps on Mac</span></span><br><span data-ttu-id="dfbc9-820">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-820">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="dfbc9-821">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-821">- Content</span></span><br><span data-ttu-id="dfbc9-822">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-822">
         - TaskPane</span></span><br><span data-ttu-id="dfbc9-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-823">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-824">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="dfbc9-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-828">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-828">- ActiveView</span></span><br><span data-ttu-id="dfbc9-829">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-829">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-830">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-830">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-831">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-831">
         - File</span></span><br><span data-ttu-id="dfbc9-832">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-832">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-833">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-833">
         - Selection</span></span><br><span data-ttu-id="dfbc9-834">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-834">
         - Settings</span></span><br><span data-ttu-id="dfbc9-835">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-835">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-836">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-836">Office 2019 for Mac</span></span><br><span data-ttu-id="dfbc9-837">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-837">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-838">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-838">- Content</span></span><br><span data-ttu-id="dfbc9-839">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-839">
         - TaskPane</span></span><br><span data-ttu-id="dfbc9-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-841">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-843">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-843">- ActiveView</span></span><br><span data-ttu-id="dfbc9-844">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-844">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-845">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-845">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-846">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-846">
         - File</span></span><br><span data-ttu-id="dfbc9-847">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-847">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-848">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-848">
         - Selection</span></span><br><span data-ttu-id="dfbc9-849">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-849">
         - Settings</span></span><br><span data-ttu-id="dfbc9-850">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-850">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-851">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-851">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="dfbc9-852">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-852">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-853">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-853">- Content</span></span><br><span data-ttu-id="dfbc9-854">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-854">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-855">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="dfbc9-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-857">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="dfbc9-857">- ActiveView</span></span><br><span data-ttu-id="dfbc9-858">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-858">
         - CompressedFile</span></span><br><span data-ttu-id="dfbc9-859">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-859">
         - DocumentEvents</span></span><br><span data-ttu-id="dfbc9-860">
         - File</span><span class="sxs-lookup"><span data-stu-id="dfbc9-860">
         - File</span></span><br><span data-ttu-id="dfbc9-861">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="dfbc9-861">
         - PdfFile</span></span><br><span data-ttu-id="dfbc9-862">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-862">
         - Selection</span></span><br><span data-ttu-id="dfbc9-863">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-863">
         - Settings</span></span><br><span data-ttu-id="dfbc9-864">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-864">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="dfbc9-865">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="dfbc9-865">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="dfbc9-866">OneNote</span><span class="sxs-lookup"><span data-stu-id="dfbc9-866">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dfbc9-867">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-867">Platform</span></span></th>
    <th><span data-ttu-id="dfbc9-868">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-868">Extension points</span></span></th>
    <th><span data-ttu-id="dfbc9-869">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-869">API requirement sets</span></span></th>
    <th><span data-ttu-id="dfbc9-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-870"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-871">Office on the web</span><span class="sxs-lookup"><span data-stu-id="dfbc9-871">Office on the web</span></span></td>
    <td> <span data-ttu-id="dfbc9-872">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-872">- Content</span></span><br><span data-ttu-id="dfbc9-873">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-873">
         - TaskPane</span></span><br><span data-ttu-id="dfbc9-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-875">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="dfbc9-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-878">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="dfbc9-878">- DocumentEvents</span></span><br><span data-ttu-id="dfbc9-879">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-879">
         - HtmlCoercion</span></span><br><span data-ttu-id="dfbc9-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="dfbc9-880">
         - Settings</span></span><br><span data-ttu-id="dfbc9-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-881">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="dfbc9-882">Project</span><span class="sxs-lookup"><span data-stu-id="dfbc9-882">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="dfbc9-883">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="dfbc9-883">Platform</span></span></th>
    <th><span data-ttu-id="dfbc9-884">拡張点</span><span class="sxs-lookup"><span data-stu-id="dfbc9-884">Extension points</span></span></th>
    <th><span data-ttu-id="dfbc9-885">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-885">API requirement sets</span></span></th>
    <th><span data-ttu-id="dfbc9-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-886"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-887">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="dfbc9-887">Office 2019 on Windows</span></span><br><span data-ttu-id="dfbc9-888">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-888">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-889">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-889">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-890">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-891">- Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-891">- Selection</span></span><br><span data-ttu-id="dfbc9-892">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-892">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-893">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="dfbc9-893">Office 2016 on Windows</span></span><br><span data-ttu-id="dfbc9-894">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-894">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-895">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-895">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-896">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-897">- Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-897">- Selection</span></span><br><span data-ttu-id="dfbc9-898">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-898">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="dfbc9-899">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="dfbc9-899">Office 2013 on Windows</span></span><br><span data-ttu-id="dfbc9-900">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="dfbc9-900">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="dfbc9-901">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="dfbc9-901">- TaskPane</span></span></td>
    <td> <span data-ttu-id="dfbc9-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="dfbc9-902">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="dfbc9-903">- Selection</span><span class="sxs-lookup"><span data-stu-id="dfbc9-903">- Selection</span></span><br><span data-ttu-id="dfbc9-904">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="dfbc9-904">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="dfbc9-905">関連項目</span><span class="sxs-lookup"><span data-stu-id="dfbc9-905">See also</span></span>

- [<span data-ttu-id="dfbc9-906">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="dfbc9-906">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="dfbc9-907">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-907">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="dfbc9-908">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-908">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="dfbc9-909">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="dfbc9-909">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="dfbc9-910">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="dfbc9-910">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="dfbc9-911">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="dfbc9-911">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="dfbc9-912">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="dfbc9-912">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="dfbc9-913">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="dfbc9-913">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="dfbc9-914">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="dfbc9-914">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="dfbc9-915">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="dfbc9-915">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="dfbc9-916">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="dfbc9-916">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
