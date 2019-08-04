---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 07/26/2019
localization_priority: Priority
ms.openlocfilehash: 7039ca59af22f1101bdff7b6bcd4506497d6c9cd
ms.sourcegitcommit: cb5e1726849aff591f19b07391198a96d5749243
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/31/2019
ms.locfileid: "35940837"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7252d-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7252d-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7252d-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="7252d-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7252d-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="7252d-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="7252d-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="7252d-108">Excel</span><span class="sxs-lookup"><span data-stu-id="7252d-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7252d-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7252d-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7252d-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7252d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="7252d-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-114">- TaskPane</span></span><br><span data-ttu-id="7252d-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-115">
        - Content</span></span><br><span data-ttu-id="7252d-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-116">
        - Custom Functions</span></span><br><span data-ttu-id="7252d-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="7252d-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7252d-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7252d-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7252d-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7252d-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7252d-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7252d-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7252d-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7252d-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7252d-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7252d-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7252d-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7252d-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-130">
        - BindingEvents</span></span><br><span data-ttu-id="7252d-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-131">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-132">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-133">
        - File</span></span><br><span data-ttu-id="7252d-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-134">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-136">
        - Selection</span></span><br><span data-ttu-id="7252d-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-137">
        - Settings</span></span><br><span data-ttu-id="7252d-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-138">
        - TableBindings</span></span><br><span data-ttu-id="7252d-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-139">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-140">
        - TextBindings</span></span><br><span data-ttu-id="7252d-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7252d-142">Office on Windows</span></span><br><span data-ttu-id="7252d-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-144">- TaskPane</span></span><br><span data-ttu-id="7252d-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-145">
        - Content</span></span><br><span data-ttu-id="7252d-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-146">
        - Custom Functions</span></span><br><span data-ttu-id="7252d-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="7252d-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7252d-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7252d-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7252d-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7252d-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7252d-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7252d-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7252d-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7252d-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7252d-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7252d-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7252d-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7252d-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-160">
        - BindingEvents</span></span><br><span data-ttu-id="7252d-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-161">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-162">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-163">
        - File</span></span><br><span data-ttu-id="7252d-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-164">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-166">
        - Selection</span></span><br><span data-ttu-id="7252d-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-167">
        - Settings</span></span><br><span data-ttu-id="7252d-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-168">
        - TableBindings</span></span><br><span data-ttu-id="7252d-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-169">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-170">
        - TextBindings</span></span><br><span data-ttu-id="7252d-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-172">Office 2019 on Windows</span></span><br><span data-ttu-id="7252d-173">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7252d-174">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-174">- TaskPane</span></span><br><span data-ttu-id="7252d-175">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-175">
        - Content</span></span><br><span data-ttu-id="7252d-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7252d-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7252d-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7252d-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7252d-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7252d-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7252d-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7252d-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7252d-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7252d-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7252d-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-187">- BindingEvents</span></span><br><span data-ttu-id="7252d-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-188">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-189">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-190">
        - File</span></span><br><span data-ttu-id="7252d-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-191">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-193">
        - Selection</span></span><br><span data-ttu-id="7252d-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-194">
        - Settings</span></span><br><span data-ttu-id="7252d-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-195">
        - TableBindings</span></span><br><span data-ttu-id="7252d-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-196">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-197">
        - TextBindings</span></span><br><span data-ttu-id="7252d-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-199">Office 2016 on Windows</span></span><br><span data-ttu-id="7252d-200">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7252d-201">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-201">- TaskPane</span></span><br><span data-ttu-id="7252d-202">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-202">
        - Content</span></span></td>
    <td><span data-ttu-id="7252d-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7252d-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7252d-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-206">- BindingEvents</span></span><br><span data-ttu-id="7252d-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-207">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-208">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-209">
        - File</span></span><br><span data-ttu-id="7252d-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-210">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-212">
        - Selection</span></span><br><span data-ttu-id="7252d-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-213">
        - Settings</span></span><br><span data-ttu-id="7252d-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-214">
        - TableBindings</span></span><br><span data-ttu-id="7252d-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-215">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-216">
        - TextBindings</span></span><br><span data-ttu-id="7252d-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7252d-218">Office 2013 on Windows</span></span><br><span data-ttu-id="7252d-219">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7252d-220">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-220">
        - TaskPane</span></span><br><span data-ttu-id="7252d-221">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7252d-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7252d-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7252d-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7252d-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-224">
        - BindingEvents</span></span><br><span data-ttu-id="7252d-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-225">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-226">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-227">
        - File</span></span><br><span data-ttu-id="7252d-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-228">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-230">
        - Selection</span></span><br><span data-ttu-id="7252d-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-231">
        - Settings</span></span><br><span data-ttu-id="7252d-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-232">
        - TableBindings</span></span><br><span data-ttu-id="7252d-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-233">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-234">
        - TextBindings</span></span><br><span data-ttu-id="7252d-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-236">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="7252d-237">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7252d-238">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-238">- TaskPane</span></span><br><span data-ttu-id="7252d-239">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-239">
        - Content</span></span><br><span data-ttu-id="7252d-240">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7252d-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7252d-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7252d-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7252d-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7252d-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7252d-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7252d-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7252d-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7252d-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7252d-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7252d-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7252d-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-252">- BindingEvents</span></span><br><span data-ttu-id="7252d-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-253">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-254">
        - File</span></span><br><span data-ttu-id="7252d-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-255">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-257">
        - Selection</span></span><br><span data-ttu-id="7252d-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-258">
        - Settings</span></span><br><span data-ttu-id="7252d-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-259">
        - TableBindings</span></span><br><span data-ttu-id="7252d-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-260">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-261">
        - TextBindings</span></span><br><span data-ttu-id="7252d-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-263">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-263">Office apps on Mac</span></span><br><span data-ttu-id="7252d-264">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7252d-265">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-265">- TaskPane</span></span><br><span data-ttu-id="7252d-266">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-266">
        - Content</span></span><br><span data-ttu-id="7252d-267">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-267">
        - Custom Functions</span></span><br><span data-ttu-id="7252d-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7252d-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7252d-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7252d-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7252d-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7252d-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7252d-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7252d-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7252d-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7252d-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7252d-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7252d-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7252d-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-281">- BindingEvents</span></span><br><span data-ttu-id="7252d-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-282">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-283">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-284">
        - File</span></span><br><span data-ttu-id="7252d-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-285">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-287">
        - PdfFile</span></span><br><span data-ttu-id="7252d-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-288">
        - Selection</span></span><br><span data-ttu-id="7252d-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-289">
        - Settings</span></span><br><span data-ttu-id="7252d-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-290">
        - TableBindings</span></span><br><span data-ttu-id="7252d-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-291">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-292">
        - TextBindings</span></span><br><span data-ttu-id="7252d-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-294">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-294">Office 2019 for Mac</span></span><br><span data-ttu-id="7252d-295">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7252d-296">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-296">- TaskPane</span></span><br><span data-ttu-id="7252d-297">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-297">
        - Content</span></span><br><span data-ttu-id="7252d-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7252d-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7252d-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7252d-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7252d-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7252d-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7252d-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7252d-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7252d-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7252d-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7252d-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-309">- BindingEvents</span></span><br><span data-ttu-id="7252d-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-310">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-311">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-312">
        - File</span></span><br><span data-ttu-id="7252d-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-313">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-315">
        - PdfFile</span></span><br><span data-ttu-id="7252d-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-316">
        - Selection</span></span><br><span data-ttu-id="7252d-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-317">
        - Settings</span></span><br><span data-ttu-id="7252d-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-318">
        - TableBindings</span></span><br><span data-ttu-id="7252d-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-319">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-320">
        - TextBindings</span></span><br><span data-ttu-id="7252d-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-322">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="7252d-323">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7252d-324">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-324">- TaskPane</span></span><br><span data-ttu-id="7252d-325">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-325">
        - Content</span></span></td>
    <td><span data-ttu-id="7252d-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7252d-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7252d-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7252d-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-329">- BindingEvents</span></span><br><span data-ttu-id="7252d-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-330">
        - CompressedFile</span></span><br><span data-ttu-id="7252d-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-331">
        - DocumentEvents</span></span><br><span data-ttu-id="7252d-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="7252d-332">
        - File</span></span><br><span data-ttu-id="7252d-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-333">
        - MatrixBindings</span></span><br><span data-ttu-id="7252d-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="7252d-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-335">
        - PdfFile</span></span><br><span data-ttu-id="7252d-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-336">
        - Selection</span></span><br><span data-ttu-id="7252d-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-337">
        - Settings</span></span><br><span data-ttu-id="7252d-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-338">
        - TableBindings</span></span><br><span data-ttu-id="7252d-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-339">
        - TableCoercion</span></span><br><span data-ttu-id="7252d-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-340">
        - TextBindings</span></span><br><span data-ttu-id="7252d-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7252d-342">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7252d-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="7252d-343">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7252d-344">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7252d-345">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7252d-346">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7252d-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-348">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-348">Office on the web</span></span></td>
    <td><span data-ttu-id="7252d-349">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7252d-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-351">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7252d-351">Office on Windows</span></span><br><span data-ttu-id="7252d-352">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7252d-353">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7252d-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="7252d-355">Office for Mac</span></span><br><span data-ttu-id="7252d-356">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="7252d-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7252d-357">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7252d-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7252d-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="7252d-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="7252d-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7252d-360">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-360">Platform</span></span></th>
    <th><span data-ttu-id="7252d-361">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-361">Extension points</span></span></th>
    <th><span data-ttu-id="7252d-362">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="7252d-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-364">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-364">Office on the web</span></span><br><span data-ttu-id="7252d-365">(モダン)</span><span class="sxs-lookup"><span data-stu-id="7252d-365">Modern</span></span></td>
    <td> <span data-ttu-id="7252d-366">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-366">- Mail Read</span></span><br><span data-ttu-id="7252d-367">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-367">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7252d-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7252d-376">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-377">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-377">Office on the web</span></span><br><span data-ttu-id="7252d-378">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="7252d-378">Classic</span></span></td>
    <td> <span data-ttu-id="7252d-379">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-379">- Mail Read</span></span><br><span data-ttu-id="7252d-380">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-380">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7252d-388">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-389">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7252d-389">Office on Windows</span></span><br><span data-ttu-id="7252d-390">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-391">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-391">- Mail Read</span></span><br><span data-ttu-id="7252d-392">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-392">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7252d-394">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="7252d-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7252d-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7252d-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7252d-402">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-403">Office 2019 on Windows</span></span><br><span data-ttu-id="7252d-404">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-405">- Mail Read</span></span><br><span data-ttu-id="7252d-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-406">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7252d-408">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="7252d-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7252d-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7252d-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7252d-416">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-417">Office 2016 on Windows</span></span><br><span data-ttu-id="7252d-418">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-419">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-419">- Mail Read</span></span><br><span data-ttu-id="7252d-420">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-420">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7252d-422">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="7252d-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7252d-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7252d-427">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7252d-428">Office 2013 on Windows</span></span><br><span data-ttu-id="7252d-429">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-430">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-430">- Mail Read</span></span><br><span data-ttu-id="7252d-431">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="7252d-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="7252d-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7252d-436">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-437">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-437">Office apps on iOS</span></span><br><span data-ttu-id="7252d-438">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-439">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-439">- Mail Read</span></span><br><span data-ttu-id="7252d-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7252d-446">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-447">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-447">Office apps on Mac</span></span><br><span data-ttu-id="7252d-448">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-449">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-449">- Mail Read</span></span><br><span data-ttu-id="7252d-450">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-450">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7252d-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7252d-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7252d-459">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-460">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-460">Office 2019 for Mac</span></span><br><span data-ttu-id="7252d-461">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-462">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-462">- Mail Read</span></span><br><span data-ttu-id="7252d-463">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-463">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7252d-471">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-472">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="7252d-473">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-474">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-474">- Mail Read</span></span><br><span data-ttu-id="7252d-475">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7252d-475">
      - Mail Compose</span></span><br><span data-ttu-id="7252d-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7252d-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7252d-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7252d-483">使用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-484">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-484">Office apps on Android</span></span><br><span data-ttu-id="7252d-485">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-486">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7252d-486">- Mail Read</span></span><br><span data-ttu-id="7252d-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7252d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7252d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7252d-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7252d-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7252d-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7252d-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7252d-493">利用不可</span><span class="sxs-lookup"><span data-stu-id="7252d-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="7252d-494">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7252d-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="7252d-495">Word</span><span class="sxs-lookup"><span data-stu-id="7252d-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7252d-496">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-496">Platform</span></span></th>
    <th><span data-ttu-id="7252d-497">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-497">Extension points</span></span></th>
    <th><span data-ttu-id="7252d-498">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="7252d-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-500">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="7252d-501">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-501">- TaskPane</span></span><br><span data-ttu-id="7252d-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7252d-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7252d-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7252d-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-509">- BindingEvents</span></span><br><span data-ttu-id="7252d-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-511">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-512">
         - File</span></span><br><span data-ttu-id="7252d-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-514">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-517">
         - PdfFile</span></span><br><span data-ttu-id="7252d-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-518">
         - Selection</span></span><br><span data-ttu-id="7252d-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-519">
         - Settings</span></span><br><span data-ttu-id="7252d-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-520">
         - TableBindings</span></span><br><span data-ttu-id="7252d-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-521">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-522">
         - TextBindings</span></span><br><span data-ttu-id="7252d-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-523">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-525">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7252d-525">Office on Windows</span></span><br><span data-ttu-id="7252d-526">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-527">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-527">- TaskPane</span></span><br><span data-ttu-id="7252d-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7252d-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7252d-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7252d-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-535">- BindingEvents</span></span><br><span data-ttu-id="7252d-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-536">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-538">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-539">
         - File</span></span><br><span data-ttu-id="7252d-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-541">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-544">
         - PdfFile</span></span><br><span data-ttu-id="7252d-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-545">
         - Selection</span></span><br><span data-ttu-id="7252d-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-546">
         - Settings</span></span><br><span data-ttu-id="7252d-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-547">
         - TableBindings</span></span><br><span data-ttu-id="7252d-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-548">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-549">
         - TextBindings</span></span><br><span data-ttu-id="7252d-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-550">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-552">Office 2019 on Windows</span></span><br><span data-ttu-id="7252d-553">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-554">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-554">- TaskPane</span></span><br><span data-ttu-id="7252d-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7252d-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7252d-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-561">- BindingEvents</span></span><br><span data-ttu-id="7252d-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-562">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-564">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-565">
         - File</span></span><br><span data-ttu-id="7252d-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-567">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-570">
         - PdfFile</span></span><br><span data-ttu-id="7252d-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-571">
         - Selection</span></span><br><span data-ttu-id="7252d-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-572">
         - Settings</span></span><br><span data-ttu-id="7252d-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-573">
         - TableBindings</span></span><br><span data-ttu-id="7252d-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-574">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-575">
         - TextBindings</span></span><br><span data-ttu-id="7252d-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-576">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-578">Office 2016 on Windows</span></span><br><span data-ttu-id="7252d-579">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-580">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7252d-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-584">- BindingEvents</span></span><br><span data-ttu-id="7252d-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-585">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-587">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-588">
         - File</span></span><br><span data-ttu-id="7252d-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-590">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-593">
         - PdfFile</span></span><br><span data-ttu-id="7252d-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-594">
         - Selection</span></span><br><span data-ttu-id="7252d-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-595">
         - Settings</span></span><br><span data-ttu-id="7252d-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-596">
         - TableBindings</span></span><br><span data-ttu-id="7252d-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-597">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-598">
         - TextBindings</span></span><br><span data-ttu-id="7252d-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-599">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7252d-601">Office 2013 on Windows</span></span><br><span data-ttu-id="7252d-602">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-603">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7252d-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7252d-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-606">- BindingEvents</span></span><br><span data-ttu-id="7252d-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-607">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-609">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-610">
         - File</span></span><br><span data-ttu-id="7252d-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-612">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-615">
         - PdfFile</span></span><br><span data-ttu-id="7252d-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-616">
         - Selection</span></span><br><span data-ttu-id="7252d-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-617">
         - Settings</span></span><br><span data-ttu-id="7252d-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-618">
         - TableBindings</span></span><br><span data-ttu-id="7252d-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-619">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-620">
         - TextBindings</span></span><br><span data-ttu-id="7252d-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-621">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-623">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="7252d-624">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-625">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7252d-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7252d-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="7252d-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-631">- BindingEvents</span></span><br><span data-ttu-id="7252d-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-632">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-634">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-635">
         - File</span></span><br><span data-ttu-id="7252d-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-637">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-640">
         - PdfFile</span></span><br><span data-ttu-id="7252d-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-641">
         - Selection</span></span><br><span data-ttu-id="7252d-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-642">
         - Settings</span></span><br><span data-ttu-id="7252d-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-643">
         - TableBindings</span></span><br><span data-ttu-id="7252d-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-644">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-645">
         - TextBindings</span></span><br><span data-ttu-id="7252d-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-646">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-648">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-648">Office apps on Mac</span></span><br><span data-ttu-id="7252d-649">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-650">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-650">- TaskPane</span></span><br><span data-ttu-id="7252d-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7252d-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7252d-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="7252d-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-658">- BindingEvents</span></span><br><span data-ttu-id="7252d-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-659">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-661">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-662">
         - File</span></span><br><span data-ttu-id="7252d-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-664">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-667">
         - PdfFile</span></span><br><span data-ttu-id="7252d-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-668">
         - Selection</span></span><br><span data-ttu-id="7252d-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-669">
         - Settings</span></span><br><span data-ttu-id="7252d-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-670">
         - TableBindings</span></span><br><span data-ttu-id="7252d-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-671">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-672">
         - TextBindings</span></span><br><span data-ttu-id="7252d-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-673">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-675">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-675">Office 2019 for Mac</span></span><br><span data-ttu-id="7252d-676">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-677">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-677">- TaskPane</span></span><br><span data-ttu-id="7252d-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7252d-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7252d-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7252d-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="7252d-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-684">- BindingEvents</span></span><br><span data-ttu-id="7252d-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-685">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-687">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-688">
         - File</span></span><br><span data-ttu-id="7252d-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-690">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-693">
         - PdfFile</span></span><br><span data-ttu-id="7252d-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-694">
         - Selection</span></span><br><span data-ttu-id="7252d-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-695">
         - Settings</span></span><br><span data-ttu-id="7252d-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-696">
         - TableBindings</span></span><br><span data-ttu-id="7252d-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-697">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-698">
         - TextBindings</span></span><br><span data-ttu-id="7252d-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-699">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-701">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="7252d-702">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-703">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7252d-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7252d-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7252d-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-707">- BindingEvents</span></span><br><span data-ttu-id="7252d-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-708">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7252d-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="7252d-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-710">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-711">
         - File</span></span><br><span data-ttu-id="7252d-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-713">
         - MatrixBindings</span></span><br><span data-ttu-id="7252d-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="7252d-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7252d-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-716">
         - PdfFile</span></span><br><span data-ttu-id="7252d-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-717">
         - Selection</span></span><br><span data-ttu-id="7252d-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-718">
         - Settings</span></span><br><span data-ttu-id="7252d-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-719">
         - TableBindings</span></span><br><span data-ttu-id="7252d-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-720">
         - TableCoercion</span></span><br><span data-ttu-id="7252d-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7252d-721">
         - TextBindings</span></span><br><span data-ttu-id="7252d-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-722">
         - TextCoercion</span></span><br><span data-ttu-id="7252d-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7252d-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="7252d-724">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7252d-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7252d-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7252d-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7252d-726">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-726">Platform</span></span></th>
    <th><span data-ttu-id="7252d-727">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-727">Extension points</span></span></th>
    <th><span data-ttu-id="7252d-728">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="7252d-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-730">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="7252d-731">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-731">- Content</span></span><br><span data-ttu-id="7252d-732">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-732">
         - TaskPane</span></span><br><span data-ttu-id="7252d-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-734">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7252d-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-737">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7252d-738">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-738">- ActiveView</span></span><br><span data-ttu-id="7252d-739">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-739">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-740">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-740">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-741">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-741">
         - File</span></span><br><span data-ttu-id="7252d-742">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-742">
         - PdfFile</span></span><br><span data-ttu-id="7252d-743">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-743">
         - Selection</span></span><br><span data-ttu-id="7252d-744">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-744">
         - Settings</span></span><br><span data-ttu-id="7252d-745">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-745">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-746">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7252d-746">Office on Windows</span></span><br><span data-ttu-id="7252d-747">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-747">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-748">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-748">- Content</span></span><br><span data-ttu-id="7252d-749">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-749">
         - TaskPane</span></span><br><span data-ttu-id="7252d-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-750">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-751">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7252d-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-754">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7252d-755">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-755">- ActiveView</span></span><br><span data-ttu-id="7252d-756">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-756">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-757">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-757">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-758">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-758">
         - File</span></span><br><span data-ttu-id="7252d-759">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-759">
         - PdfFile</span></span><br><span data-ttu-id="7252d-760">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-760">
         - Selection</span></span><br><span data-ttu-id="7252d-761">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-761">
         - Settings</span></span><br><span data-ttu-id="7252d-762">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-762">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-763">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-763">Office 2019 on Windows</span></span><br><span data-ttu-id="7252d-764">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-764">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-765">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-765">- Content</span></span><br><span data-ttu-id="7252d-766">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-766">
         - TaskPane</span></span><br><span data-ttu-id="7252d-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-768">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-769">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-770">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-770">- ActiveView</span></span><br><span data-ttu-id="7252d-771">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-771">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-772">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-772">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-773">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-773">
         - File</span></span><br><span data-ttu-id="7252d-774">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-774">
         - PdfFile</span></span><br><span data-ttu-id="7252d-775">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-775">
         - Selection</span></span><br><span data-ttu-id="7252d-776">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-776">
         - Settings</span></span><br><span data-ttu-id="7252d-777">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-777">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-778">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-778">Office 2016 on Windows</span></span><br><span data-ttu-id="7252d-779">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-779">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-780">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-780">- Content</span></span><br><span data-ttu-id="7252d-781">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-781">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7252d-782">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7252d-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-784">- ActiveView</span></span><br><span data-ttu-id="7252d-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-785">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-786">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-787">
         - File</span></span><br><span data-ttu-id="7252d-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-788">
         - PdfFile</span></span><br><span data-ttu-id="7252d-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-789">
         - Selection</span></span><br><span data-ttu-id="7252d-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-790">
         - Settings</span></span><br><span data-ttu-id="7252d-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-792">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7252d-792">Office 2013 on Windows</span></span><br><span data-ttu-id="7252d-793">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-794">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-794">- Content</span></span><br><span data-ttu-id="7252d-795">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-795">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="7252d-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7252d-796">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7252d-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-797">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-798">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-798">- ActiveView</span></span><br><span data-ttu-id="7252d-799">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-799">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-800">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-800">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-801">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-801">
         - File</span></span><br><span data-ttu-id="7252d-802">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-802">
         - PdfFile</span></span><br><span data-ttu-id="7252d-803">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-803">
         - Selection</span></span><br><span data-ttu-id="7252d-804">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-804">
         - Settings</span></span><br><span data-ttu-id="7252d-805">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-805">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-806">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-806">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="7252d-807">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-807">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-808">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-808">- Content</span></span><br><span data-ttu-id="7252d-809">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-809">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-810">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7252d-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-813">- ActiveView</span></span><br><span data-ttu-id="7252d-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-814">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-815">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-816">
         - File</span></span><br><span data-ttu-id="7252d-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-817">
         - PdfFile</span></span><br><span data-ttu-id="7252d-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-818">
         - Selection</span></span><br><span data-ttu-id="7252d-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-819">
         - Settings</span></span><br><span data-ttu-id="7252d-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-821">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7252d-821">Office apps on Mac</span></span><br><span data-ttu-id="7252d-822">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7252d-822">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7252d-823">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-823">- Content</span></span><br><span data-ttu-id="7252d-824">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-824">
         - TaskPane</span></span><br><span data-ttu-id="7252d-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-826">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7252d-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7252d-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7252d-829">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7252d-830">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-830">- ActiveView</span></span><br><span data-ttu-id="7252d-831">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-831">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-832">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-832">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-833">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-833">
         - File</span></span><br><span data-ttu-id="7252d-834">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-834">
         - PdfFile</span></span><br><span data-ttu-id="7252d-835">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-835">
         - Selection</span></span><br><span data-ttu-id="7252d-836">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-836">
         - Settings</span></span><br><span data-ttu-id="7252d-837">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-837">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-838">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-838">Office 2019 for Mac</span></span><br><span data-ttu-id="7252d-839">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-839">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-840">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-840">- Content</span></span><br><span data-ttu-id="7252d-841">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-841">
         - TaskPane</span></span><br><span data-ttu-id="7252d-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-842">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-843">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-844">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-845">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-845">- ActiveView</span></span><br><span data-ttu-id="7252d-846">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-846">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-847">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-847">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-848">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-848">
         - File</span></span><br><span data-ttu-id="7252d-849">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-849">
         - PdfFile</span></span><br><span data-ttu-id="7252d-850">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-850">
         - Selection</span></span><br><span data-ttu-id="7252d-851">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-851">
         - Settings</span></span><br><span data-ttu-id="7252d-852">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-852">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-853">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-853">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="7252d-854">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-854">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-855">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-855">- Content</span></span><br><span data-ttu-id="7252d-856">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-856">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7252d-857">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7252d-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7252d-859">- ActiveView</span></span><br><span data-ttu-id="7252d-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7252d-860">
         - CompressedFile</span></span><br><span data-ttu-id="7252d-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-861">
         - DocumentEvents</span></span><br><span data-ttu-id="7252d-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="7252d-862">
         - File</span></span><br><span data-ttu-id="7252d-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7252d-863">
         - PdfFile</span></span><br><span data-ttu-id="7252d-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-864">
         - Selection</span></span><br><span data-ttu-id="7252d-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-865">
         - Settings</span></span><br><span data-ttu-id="7252d-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-866">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7252d-867">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7252d-867">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="7252d-868">OneNote</span><span class="sxs-lookup"><span data-stu-id="7252d-868">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7252d-869">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-869">Platform</span></span></th>
    <th><span data-ttu-id="7252d-870">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-870">Extension points</span></span></th>
    <th><span data-ttu-id="7252d-871">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-871">API requirement sets</span></span></th>
    <th><span data-ttu-id="7252d-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-872"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-873">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7252d-873">Office on the web</span></span></td>
    <td> <span data-ttu-id="7252d-874">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7252d-874">- Content</span></span><br><span data-ttu-id="7252d-875">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-875">
         - TaskPane</span></span><br><span data-ttu-id="7252d-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7252d-876">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7252d-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-877">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7252d-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7252d-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-879">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-880">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7252d-880">- DocumentEvents</span></span><br><span data-ttu-id="7252d-881">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-881">
         - HtmlCoercion</span></span><br><span data-ttu-id="7252d-882">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7252d-882">
         - Settings</span></span><br><span data-ttu-id="7252d-883">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-883">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="7252d-884">Project</span><span class="sxs-lookup"><span data-stu-id="7252d-884">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7252d-885">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7252d-885">Platform</span></span></th>
    <th><span data-ttu-id="7252d-886">拡張点</span><span class="sxs-lookup"><span data-stu-id="7252d-886">Extension points</span></span></th>
    <th><span data-ttu-id="7252d-887">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-887">API requirement sets</span></span></th>
    <th><span data-ttu-id="7252d-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7252d-888"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-889">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7252d-889">Office 2019 on Windows</span></span><br><span data-ttu-id="7252d-890">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-890">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-891">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-891">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-892">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-893">- Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-893">- Selection</span></span><br><span data-ttu-id="7252d-894">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-894">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-895">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7252d-895">Office 2016 on Windows</span></span><br><span data-ttu-id="7252d-896">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-896">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-897">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-897">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-898">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-899">- Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-899">- Selection</span></span><br><span data-ttu-id="7252d-900">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-900">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7252d-901">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7252d-901">Office 2013 on Windows</span></span><br><span data-ttu-id="7252d-902">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7252d-902">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7252d-903">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7252d-903">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7252d-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7252d-904">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7252d-905">- Selection</span><span class="sxs-lookup"><span data-stu-id="7252d-905">- Selection</span></span><br><span data-ttu-id="7252d-906">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7252d-906">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="7252d-907">関連項目</span><span class="sxs-lookup"><span data-stu-id="7252d-907">See also</span></span>

- [<span data-ttu-id="7252d-908">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="7252d-908">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7252d-909">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-909">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="7252d-910">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-910">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="7252d-911">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="7252d-911">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="7252d-912">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="7252d-912">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="7252d-913">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="7252d-913">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="7252d-914">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="7252d-914">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="7252d-915">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="7252d-915">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="7252d-916">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="7252d-916">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="7252d-917">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="7252d-917">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="7252d-918">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="7252d-918">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
