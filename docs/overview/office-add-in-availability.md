---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 04/07/2020
localization_priority: Priority
ms.openlocfilehash: 823fd53e71c71f4a845f9a7b5c6177ad3f14745f
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185618"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ec48f-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ec48f-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ec48f-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ec48f-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ec48f-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ec48f-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="ec48f-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ec48f-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ec48f-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ec48f-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ec48f-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ec48f-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ec48f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec48f-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-114">- TaskPane</span></span><br><span data-ttu-id="ec48f-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-115">
        - Content</span></span><br><span data-ttu-id="ec48f-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ec48f-116">
        - Custom Functions</span></span><br><span data-ttu-id="ec48f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="ec48f-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ec48f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec48f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec48f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec48f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec48f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec48f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec48f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec48f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec48f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ec48f-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="ec48f-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-130">
        - BindingEvents</span></span><br><span data-ttu-id="ec48f-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-131">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-132">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-133">
        - File</span></span><br><span data-ttu-id="ec48f-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-134">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-136">
        - Selection</span></span><br><span data-ttu-id="ec48f-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-137">
        - Settings</span></span><br><span data-ttu-id="ec48f-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-138">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-139">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-140">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-142">Office on Windows</span></span><br><span data-ttu-id="ec48f-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-144">- TaskPane</span></span><br><span data-ttu-id="ec48f-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-145">
        - Content</span></span><br><span data-ttu-id="ec48f-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ec48f-146">
        - Custom Functions</span></span><br><span data-ttu-id="ec48f-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="ec48f-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ec48f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec48f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec48f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec48f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec48f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec48f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec48f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec48f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec48f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ec48f-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ec48f-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-161">
        - BindingEvents</span></span><br><span data-ttu-id="ec48f-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-162">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-163">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-164">
        - File</span></span><br><span data-ttu-id="ec48f-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-165">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-167">
        - Selection</span></span><br><span data-ttu-id="ec48f-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-168">
        - Settings</span></span><br><span data-ttu-id="ec48f-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-169">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-170">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-171">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-173">Office 2019 on Windows</span></span><br><span data-ttu-id="ec48f-174">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec48f-175">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-175">- TaskPane</span></span><br><span data-ttu-id="ec48f-176">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-176">
        - Content</span></span><br><span data-ttu-id="ec48f-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ec48f-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec48f-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec48f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec48f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec48f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec48f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec48f-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec48f-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-188">- BindingEvents</span></span><br><span data-ttu-id="ec48f-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-189">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-190">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-191">
        - File</span></span><br><span data-ttu-id="ec48f-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-192">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-194">
        - Selection</span></span><br><span data-ttu-id="ec48f-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-195">
        - Settings</span></span><br><span data-ttu-id="ec48f-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-196">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-197">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-198">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-200">Office 2016 on Windows</span></span><br><span data-ttu-id="ec48f-201">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec48f-202">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-202">- TaskPane</span></span><br><span data-ttu-id="ec48f-203">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-203">
        - Content</span></span></td>
    <td><span data-ttu-id="ec48f-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec48f-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-207">- BindingEvents</span></span><br><span data-ttu-id="ec48f-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-208">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-209">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-210">
        - File</span></span><br><span data-ttu-id="ec48f-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-211">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-213">
        - Selection</span></span><br><span data-ttu-id="ec48f-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-214">
        - Settings</span></span><br><span data-ttu-id="ec48f-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-215">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-216">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-217">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ec48f-219">Office 2013 on Windows</span></span><br><span data-ttu-id="ec48f-220">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec48f-221">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-221">
        - TaskPane</span></span><br><span data-ttu-id="ec48f-222">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ec48f-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec48f-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec48f-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-225">
        - BindingEvents</span></span><br><span data-ttu-id="ec48f-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-226">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-227">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-228">
        - File</span></span><br><span data-ttu-id="ec48f-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-229">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-231">
        - Selection</span></span><br><span data-ttu-id="ec48f-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-232">
        - Settings</span></span><br><span data-ttu-id="ec48f-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-233">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-234">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-235">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-237">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-237">Office on iPad</span></span><br><span data-ttu-id="ec48f-238">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ec48f-239">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-239">- TaskPane</span></span><br><span data-ttu-id="ec48f-240">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-240">
        - Content</span></span></td>
    <td><span data-ttu-id="ec48f-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec48f-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec48f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec48f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec48f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec48f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec48f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec48f-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec48f-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ec48f-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-253">- BindingEvents</span></span><br><span data-ttu-id="ec48f-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-254">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-255">
        - File</span></span><br><span data-ttu-id="ec48f-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-256">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-258">
        - Selection</span></span><br><span data-ttu-id="ec48f-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-259">
        - Settings</span></span><br><span data-ttu-id="ec48f-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-260">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-261">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-262">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-264">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-264">Office on Mac</span></span><br><span data-ttu-id="ec48f-265">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ec48f-266">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-266">- TaskPane</span></span><br><span data-ttu-id="ec48f-267">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-267">
        - Content</span></span><br><span data-ttu-id="ec48f-268">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ec48f-268">
        - Custom Functions</span></span><br><span data-ttu-id="ec48f-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ec48f-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec48f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec48f-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec48f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec48f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec48f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec48f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec48f-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ec48f-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ec48f-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ec48f-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-283">- BindingEvents</span></span><br><span data-ttu-id="ec48f-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-284">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-285">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-286">
        - File</span></span><br><span data-ttu-id="ec48f-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-287">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-289">
        - PdfFile</span></span><br><span data-ttu-id="ec48f-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-290">
        - Selection</span></span><br><span data-ttu-id="ec48f-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-291">
        - Settings</span></span><br><span data-ttu-id="ec48f-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-292">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-293">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-294">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-296">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-296">Office 2019 on Mac</span></span><br><span data-ttu-id="ec48f-297">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec48f-298">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-298">- TaskPane</span></span><br><span data-ttu-id="ec48f-299">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-299">
        - Content</span></span><br><span data-ttu-id="ec48f-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ec48f-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ec48f-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ec48f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ec48f-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ec48f-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ec48f-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ec48f-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ec48f-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-311">- BindingEvents</span></span><br><span data-ttu-id="ec48f-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-312">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-313">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-314">
        - File</span></span><br><span data-ttu-id="ec48f-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-315">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-317">
        - PdfFile</span></span><br><span data-ttu-id="ec48f-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-318">
        - Selection</span></span><br><span data-ttu-id="ec48f-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-319">
        - Settings</span></span><br><span data-ttu-id="ec48f-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-320">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-321">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-322">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-324">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-324">Office 2016 on Mac</span></span><br><span data-ttu-id="ec48f-325">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ec48f-326">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-326">- TaskPane</span></span><br><span data-ttu-id="ec48f-327">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-327">
        - Content</span></span></td>
    <td><span data-ttu-id="ec48f-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ec48f-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec48f-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ec48f-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-331">- BindingEvents</span></span><br><span data-ttu-id="ec48f-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-332">
        - CompressedFile</span></span><br><span data-ttu-id="ec48f-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-333">
        - DocumentEvents</span></span><br><span data-ttu-id="ec48f-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-334">
        - File</span></span><br><span data-ttu-id="ec48f-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-335">
        - MatrixBindings</span></span><br><span data-ttu-id="ec48f-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-337">
        - PdfFile</span></span><br><span data-ttu-id="ec48f-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-338">
        - Selection</span></span><br><span data-ttu-id="ec48f-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-339">
        - Settings</span></span><br><span data-ttu-id="ec48f-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-340">
        - TableBindings</span></span><br><span data-ttu-id="ec48f-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-341">
        - TableCoercion</span></span><br><span data-ttu-id="ec48f-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-342">
        - TextBindings</span></span><br><span data-ttu-id="ec48f-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ec48f-344">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ec48f-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ec48f-345">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="ec48f-345">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ec48f-346">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ec48f-347">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ec48f-348">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ec48f-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-350">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-350">Office on the web</span></span></td>
    <td><span data-ttu-id="ec48f-351">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ec48f-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec48f-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-352">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-353">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-353">Office on Windows</span></span><br><span data-ttu-id="ec48f-354">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ec48f-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ec48f-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec48f-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-356">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="ec48f-357">Office for Mac</span></span><br><span data-ttu-id="ec48f-358">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="ec48f-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ec48f-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ec48f-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ec48f-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-360">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ec48f-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="ec48f-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec48f-362">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-362">Platform</span></span></th>
    <th><span data-ttu-id="ec48f-363">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-363">Extension points</span></span></th>
    <th><span data-ttu-id="ec48f-364">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec48f-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-366">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-366">Office on the web</span></span><br><span data-ttu-id="ec48f-367">(モダン)</span><span class="sxs-lookup"><span data-stu-id="ec48f-367">(modern)</span></span></td>
    <td> <span data-ttu-id="ec48f-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-368">- Message Read</span></span><br><span data-ttu-id="ec48f-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-369">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-370">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-370">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-371">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-371">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-372">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-373">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec48f-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ec48f-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ec48f-381">利用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-381">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-382">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-382">Office on the web</span></span><br><span data-ttu-id="ec48f-383">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="ec48f-383">(classic)</span></span></td>
    <td> <span data-ttu-id="ec48f-384">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-384">- Message Read</span></span><br><span data-ttu-id="ec48f-385">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-385">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-386">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-386">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-387">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-387">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-388">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-389">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ec48f-395">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-395">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-396">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-396">Office on Windows</span></span><br><span data-ttu-id="ec48f-397">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="ec48f-397">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-398">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-398">- Message Read</span></span><br><span data-ttu-id="ec48f-399">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-399">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-400">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-400">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-401">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-401">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-402">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ec48f-403">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="ec48f-403">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ec48f-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-404">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec48f-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ec48f-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ec48f-412">利用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-412">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-413">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-413">Office 2019 on Windows</span></span><br><span data-ttu-id="ec48f-414">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-414">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-415">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-415">- Message Read</span></span><br><span data-ttu-id="ec48f-416">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-416">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-417">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-417">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-418">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-418">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-419">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ec48f-420">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="ec48f-420">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ec48f-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-421">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec48f-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ec48f-428">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-428">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-429">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-429">Office 2016 on Windows</span></span><br><span data-ttu-id="ec48f-430">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-430">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-431">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-431">- Message Read</span></span><br><span data-ttu-id="ec48f-432">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-432">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-433">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-433">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-434">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-434">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-435">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ec48f-436">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="ec48f-436">
      - Modules</span></span></td>
    <td> <span data-ttu-id="ec48f-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-437">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ec48f-441">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-441">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-442">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ec48f-442">Office 2013 on Windows</span></span><br><span data-ttu-id="ec48f-443">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-443">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-444">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-444">- Message Read</span></span><br><span data-ttu-id="ec48f-445">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-445">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-446">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-446">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-447">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-447">
      - Appointment Organizer (Compose)</span></span><br>
    <td> <span data-ttu-id="ec48f-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-448">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ec48f-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ec48f-452">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-452">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-453">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-453">Office on iOS</span></span><br><span data-ttu-id="ec48f-454">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="ec48f-454">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-455">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-455">- Message Read</span></span><br><span data-ttu-id="ec48f-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-456">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-457">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ec48f-462">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-462">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-463">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-463">Office on Mac</span></span><br><span data-ttu-id="ec48f-464">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="ec48f-464">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-465">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-465">- Message Read</span></span><br><span data-ttu-id="ec48f-466">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-466">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-467">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-467">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-468">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-468">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ec48f-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ec48f-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ec48f-478">利用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-478">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-479">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-479">Office 2019 on Mac</span></span><br><span data-ttu-id="ec48f-480">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-480">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-481">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-481">- Message Read</span></span><br><span data-ttu-id="ec48f-482">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-482">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-483">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-483">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-484">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-484">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-485">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-486">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ec48f-492">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-492">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-493">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-493">Office 2016 on Mac</span></span><br><span data-ttu-id="ec48f-494">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-494">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-495">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-495">- Message Read</span></span><br><span data-ttu-id="ec48f-496">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="ec48f-496">
      - Message Compose</span></span><br><span data-ttu-id="ec48f-497">
      - 予定の出席者 (読み取り)</span><span class="sxs-lookup"><span data-stu-id="ec48f-497">
      - Appointment Attendee (Read)</span></span><br><span data-ttu-id="ec48f-498">
      - 予定の開催者 (作成)</span><span class="sxs-lookup"><span data-stu-id="ec48f-498">
      - Appointment Organizer (Compose)</span></span><br><span data-ttu-id="ec48f-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-499">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-500">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-501">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-502">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ec48f-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ec48f-506">使用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-506">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-507">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-507">Office on Android</span></span><br><span data-ttu-id="ec48f-508">(Office 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="ec48f-508">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-509">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="ec48f-509">- Message Read</span></span><br><span data-ttu-id="ec48f-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-510">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-511">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ec48f-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-512">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ec48f-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-513">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ec48f-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-514">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ec48f-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-515">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ec48f-516">利用不可</span><span class="sxs-lookup"><span data-stu-id="ec48f-516">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ec48f-517">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ec48f-517">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ec48f-518">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="ec48f-518">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ec48f-519">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ec48f-519">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ec48f-520">Word</span><span class="sxs-lookup"><span data-stu-id="ec48f-520">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec48f-521">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-521">Platform</span></span></th>
    <th><span data-ttu-id="ec48f-522">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-522">Extension points</span></span></th>
    <th><span data-ttu-id="ec48f-523">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-523">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec48f-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-524"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-525">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-525">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec48f-526">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-526">- TaskPane</span></span><br><span data-ttu-id="ec48f-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-529">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec48f-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-530">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec48f-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-532">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec48f-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-534">- BindingEvents</span></span><br><span data-ttu-id="ec48f-535">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-535">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-536">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-536">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-537">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-537">
         - File</span></span><br><span data-ttu-id="ec48f-538">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-538">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-539">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-539">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-540">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-540">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-541">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-541">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-542">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-542">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-543">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-543">
         - Selection</span></span><br><span data-ttu-id="ec48f-544">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-544">
         - Settings</span></span><br><span data-ttu-id="ec48f-545">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-545">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-546">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-546">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-547">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-547">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-548">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-548">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-549">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-549">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-550">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-550">Office on Windows</span></span><br><span data-ttu-id="ec48f-551">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-551">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-552">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-552">- TaskPane</span></span><br><span data-ttu-id="ec48f-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-553">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-554">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-555">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec48f-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec48f-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-558">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec48f-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-560">- BindingEvents</span></span><br><span data-ttu-id="ec48f-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-561">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-563">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-564">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-564">
         - File</span></span><br><span data-ttu-id="ec48f-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-566">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-569">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-570">
         - Selection</span></span><br><span data-ttu-id="ec48f-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-571">
         - Settings</span></span><br><span data-ttu-id="ec48f-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-572">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-573">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-574">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-575">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-577">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-577">Office 2019 on Windows</span></span><br><span data-ttu-id="ec48f-578">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-579">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-579">- TaskPane</span></span><br><span data-ttu-id="ec48f-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-580">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-582">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec48f-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-583">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec48f-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-584">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-585">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-586">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-586">- BindingEvents</span></span><br><span data-ttu-id="ec48f-587">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-587">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-588">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-588">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-589">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-589">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-590">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-590">
         - File</span></span><br><span data-ttu-id="ec48f-591">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-591">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-592">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-592">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-593">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-593">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-594">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-594">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-595">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-595">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-596">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-596">
         - Selection</span></span><br><span data-ttu-id="ec48f-597">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-597">
         - Settings</span></span><br><span data-ttu-id="ec48f-598">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-598">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-599">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-599">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-600">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-600">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-601">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-601">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-602">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-602">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-603">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-603">Office 2016 on Windows</span></span><br><span data-ttu-id="ec48f-604">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-604">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-605">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-605">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-606">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-607">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec48f-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-608">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-609">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-609">- BindingEvents</span></span><br><span data-ttu-id="ec48f-610">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-610">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-611">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-611">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-612">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-612">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-613">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-613">
         - File</span></span><br><span data-ttu-id="ec48f-614">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-614">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-615">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-615">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-616">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-616">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-617">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-617">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-618">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-618">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-619">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-619">
         - Selection</span></span><br><span data-ttu-id="ec48f-620">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-620">
         - Settings</span></span><br><span data-ttu-id="ec48f-621">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-621">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-622">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-622">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-623">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-623">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-624">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-624">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-625">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-625">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-626">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ec48f-626">Office 2013 on Windows</span></span><br><span data-ttu-id="ec48f-627">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-627">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-628">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-628">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec48f-629">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec48f-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-630">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-631">- BindingEvents</span></span><br><span data-ttu-id="ec48f-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-632">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-634">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-635">
         - File</span></span><br><span data-ttu-id="ec48f-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-637">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-640">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-641">
         - Selection</span></span><br><span data-ttu-id="ec48f-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-642">
         - Settings</span></span><br><span data-ttu-id="ec48f-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-643">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-644">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-645">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-646">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-647">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-648">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-648">Office on iPad</span></span><br><span data-ttu-id="ec48f-649">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-650">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-650">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-652">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec48f-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-653">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec48f-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-654">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-655">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ec48f-656">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-656">- BindingEvents</span></span><br><span data-ttu-id="ec48f-657">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-657">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-658">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-658">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-659">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-659">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-660">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-660">
         - File</span></span><br><span data-ttu-id="ec48f-661">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-661">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-662">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-662">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-663">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-663">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-664">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-664">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-665">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-665">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-666">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-666">
         - Selection</span></span><br><span data-ttu-id="ec48f-667">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-667">
         - Settings</span></span><br><span data-ttu-id="ec48f-668">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-668">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-669">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-669">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-670">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-670">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-671">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-671">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-672">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-672">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-673">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-673">Office on Mac</span></span><br><span data-ttu-id="ec48f-674">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-674">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-675">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-675">- TaskPane</span></span><br><span data-ttu-id="ec48f-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-676">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-677">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-678">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec48f-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec48f-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ec48f-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-683">- BindingEvents</span></span><br><span data-ttu-id="ec48f-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-684">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-686">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-687">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-687">
         - File</span></span><br><span data-ttu-id="ec48f-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-689">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-692">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-693">
         - Selection</span></span><br><span data-ttu-id="ec48f-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-694">
         - Settings</span></span><br><span data-ttu-id="ec48f-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-695">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-696">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-697">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-698">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-700">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-700">Office 2019 on Mac</span></span><br><span data-ttu-id="ec48f-701">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-702">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-702">- TaskPane</span></span><br><span data-ttu-id="ec48f-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-703">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ec48f-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-706">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ec48f-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-707">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-708">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ec48f-709">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-709">- BindingEvents</span></span><br><span data-ttu-id="ec48f-710">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-710">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-711">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-711">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-712">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-712">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-713">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-713">
         - File</span></span><br><span data-ttu-id="ec48f-714">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-714">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-715">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-715">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-716">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-716">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-717">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-717">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-718">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-718">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-719">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-719">
         - Selection</span></span><br><span data-ttu-id="ec48f-720">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-720">
         - Settings</span></span><br><span data-ttu-id="ec48f-721">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-721">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-722">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-722">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-723">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-723">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-724">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-724">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-725">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-725">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-726">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-726">Office 2016 on Mac</span></span><br><span data-ttu-id="ec48f-727">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-727">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-728">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-728">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-729">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ec48f-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ec48f-730">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ec48f-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-731">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-732">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-732">- BindingEvents</span></span><br><span data-ttu-id="ec48f-733">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-733">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-734">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ec48f-734">
         - CustomXmlParts</span></span><br><span data-ttu-id="ec48f-735">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-735">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-736">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-736">
         - File</span></span><br><span data-ttu-id="ec48f-737">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-737">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-738">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-738">
         - MatrixBindings</span></span><br><span data-ttu-id="ec48f-739">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-739">
         - MatrixCoercion</span></span><br><span data-ttu-id="ec48f-740">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-740">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ec48f-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-741">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-742">
         - Selection</span></span><br><span data-ttu-id="ec48f-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-743">
         - Settings</span></span><br><span data-ttu-id="ec48f-744">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-744">
         - TableBindings</span></span><br><span data-ttu-id="ec48f-745">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-745">
         - TableCoercion</span></span><br><span data-ttu-id="ec48f-746">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ec48f-746">
         - TextBindings</span></span><br><span data-ttu-id="ec48f-747">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-747">
         - TextCoercion</span></span><br><span data-ttu-id="ec48f-748">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-748">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ec48f-749">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ec48f-749">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ec48f-750">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ec48f-750">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec48f-751">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-751">Platform</span></span></th>
    <th><span data-ttu-id="ec48f-752">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-752">Extension points</span></span></th>
    <th><span data-ttu-id="ec48f-753">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-753">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec48f-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-754"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-755">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-755">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec48f-756">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-756">- Content</span></span><br><span data-ttu-id="ec48f-757">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-757">
         - TaskPane</span></span><br><span data-ttu-id="ec48f-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-758">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-759">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec48f-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec48f-763">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-763">- ActiveView</span></span><br><span data-ttu-id="ec48f-764">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-764">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-765">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-765">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-766">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-766">
         - File</span></span><br><span data-ttu-id="ec48f-767">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-767">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-768">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-768">
         - Selection</span></span><br><span data-ttu-id="ec48f-769">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-769">
         - Settings</span></span><br><span data-ttu-id="ec48f-770">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-770">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-771">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-771">Office on Windows</span></span><br><span data-ttu-id="ec48f-772">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-772">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-773">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-773">- Content</span></span><br><span data-ttu-id="ec48f-774">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-774">
         - TaskPane</span></span><br><span data-ttu-id="ec48f-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-775">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-776">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec48f-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-777">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-778">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec48f-780">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-780">- ActiveView</span></span><br><span data-ttu-id="ec48f-781">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-781">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-782">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-782">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-783">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-783">
         - File</span></span><br><span data-ttu-id="ec48f-784">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-784">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-785">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-785">
         - Selection</span></span><br><span data-ttu-id="ec48f-786">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-786">
         - Settings</span></span><br><span data-ttu-id="ec48f-787">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-787">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-788">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-788">Office 2019 on Windows</span></span><br><span data-ttu-id="ec48f-789">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-789">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-790">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-790">- Content</span></span><br><span data-ttu-id="ec48f-791">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-791">
         - TaskPane</span></span><br><span data-ttu-id="ec48f-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-792">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-793">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-794">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-795">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-795">- ActiveView</span></span><br><span data-ttu-id="ec48f-796">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-796">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-797">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-797">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-798">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-798">
         - File</span></span><br><span data-ttu-id="ec48f-799">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-799">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-800">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-800">
         - Selection</span></span><br><span data-ttu-id="ec48f-801">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-801">
         - Settings</span></span><br><span data-ttu-id="ec48f-802">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-802">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-803">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-803">Office 2016 on Windows</span></span><br><span data-ttu-id="ec48f-804">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-804">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-805">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-805">- Content</span></span><br><span data-ttu-id="ec48f-806">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-806">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec48f-807">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec48f-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-808">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-809">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-809">- ActiveView</span></span><br><span data-ttu-id="ec48f-810">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-810">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-811">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-811">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-812">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-812">
         - File</span></span><br><span data-ttu-id="ec48f-813">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-813">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-814">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-814">
         - Selection</span></span><br><span data-ttu-id="ec48f-815">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-815">
         - Settings</span></span><br><span data-ttu-id="ec48f-816">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-816">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-817">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ec48f-817">Office 2013 on Windows</span></span><br><span data-ttu-id="ec48f-818">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-818">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-819">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-819">- Content</span></span><br><span data-ttu-id="ec48f-820">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-820">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ec48f-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec48f-821">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec48f-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-823">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-823">- ActiveView</span></span><br><span data-ttu-id="ec48f-824">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-824">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-825">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-825">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-826">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-826">
         - File</span></span><br><span data-ttu-id="ec48f-827">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-827">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-828">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-828">
         - Selection</span></span><br><span data-ttu-id="ec48f-829">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-829">
         - Settings</span></span><br><span data-ttu-id="ec48f-830">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-830">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-831">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-831">Office on iPad</span></span><br><span data-ttu-id="ec48f-832">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-832">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-833">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-833">- Content</span></span><br><span data-ttu-id="ec48f-834">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-834">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-835">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec48f-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-837">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-838">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-838">- ActiveView</span></span><br><span data-ttu-id="ec48f-839">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-839">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-840">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-840">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-841">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-841">
         - File</span></span><br><span data-ttu-id="ec48f-842">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-842">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-843">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-843">
         - Selection</span></span><br><span data-ttu-id="ec48f-844">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-844">
         - Settings</span></span><br><span data-ttu-id="ec48f-845">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-845">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-846">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ec48f-846">Office on Mac</span></span><br><span data-ttu-id="ec48f-847">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ec48f-847">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ec48f-848">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-848">- Content</span></span><br><span data-ttu-id="ec48f-849">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-849">
         - TaskPane</span></span><br><span data-ttu-id="ec48f-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-850">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-851">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ec48f-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-852">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-853">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ec48f-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ec48f-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-855">- ActiveView</span></span><br><span data-ttu-id="ec48f-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-856">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-857">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-858">
         - File</span></span><br><span data-ttu-id="ec48f-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-859">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-860">
         - Selection</span></span><br><span data-ttu-id="ec48f-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-861">
         - Settings</span></span><br><span data-ttu-id="ec48f-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-862">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-863">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-863">Office 2019 on Mac</span></span><br><span data-ttu-id="ec48f-864">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-864">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-865">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-865">- Content</span></span><br><span data-ttu-id="ec48f-866">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-866">
         - TaskPane</span></span><br><span data-ttu-id="ec48f-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-867">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-868">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-869">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-870">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-870">- ActiveView</span></span><br><span data-ttu-id="ec48f-871">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-871">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-872">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-872">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-873">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-873">
         - File</span></span><br><span data-ttu-id="ec48f-874">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-874">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-875">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-875">
         - Selection</span></span><br><span data-ttu-id="ec48f-876">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-876">
         - Settings</span></span><br><span data-ttu-id="ec48f-877">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-877">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-878">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-878">Office 2016 on Mac</span></span><br><span data-ttu-id="ec48f-879">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-879">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-880">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-880">- Content</span></span><br><span data-ttu-id="ec48f-881">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-881">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ec48f-882">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ec48f-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-884">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ec48f-884">- ActiveView</span></span><br><span data-ttu-id="ec48f-885">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-885">
         - CompressedFile</span></span><br><span data-ttu-id="ec48f-886">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-886">
         - DocumentEvents</span></span><br><span data-ttu-id="ec48f-887">
         - File</span><span class="sxs-lookup"><span data-stu-id="ec48f-887">
         - File</span></span><br><span data-ttu-id="ec48f-888">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ec48f-888">
         - PdfFile</span></span><br><span data-ttu-id="ec48f-889">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-889">
         - Selection</span></span><br><span data-ttu-id="ec48f-890">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-890">
         - Settings</span></span><br><span data-ttu-id="ec48f-891">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-891">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ec48f-892">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ec48f-892">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ec48f-893">OneNote</span><span class="sxs-lookup"><span data-stu-id="ec48f-893">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec48f-894">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-894">Platform</span></span></th>
    <th><span data-ttu-id="ec48f-895">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-895">Extension points</span></span></th>
    <th><span data-ttu-id="ec48f-896">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-896">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec48f-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-897"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-898">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ec48f-898">Office on the web</span></span></td>
    <td> <span data-ttu-id="ec48f-899">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ec48f-899">- Content</span></span><br><span data-ttu-id="ec48f-900">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-900">
         - TaskPane</span></span><br><span data-ttu-id="ec48f-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-901">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ec48f-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-902">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ec48f-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-903">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ec48f-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-904">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-905">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ec48f-905">- DocumentEvents</span></span><br><span data-ttu-id="ec48f-906">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-906">
         - HtmlCoercion</span></span><br><span data-ttu-id="ec48f-907">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ec48f-907">
         - Settings</span></span><br><span data-ttu-id="ec48f-908">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-908">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ec48f-909">Project</span><span class="sxs-lookup"><span data-stu-id="ec48f-909">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ec48f-910">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ec48f-910">Platform</span></span></th>
    <th><span data-ttu-id="ec48f-911">拡張点</span><span class="sxs-lookup"><span data-stu-id="ec48f-911">Extension points</span></span></th>
    <th><span data-ttu-id="ec48f-912">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-912">API requirement sets</span></span></th>
    <th><span data-ttu-id="ec48f-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ec48f-913"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-914">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ec48f-914">Office 2019 on Windows</span></span><br><span data-ttu-id="ec48f-915">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-915">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-916">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-916">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-917">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-918">- Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-918">- Selection</span></span><br><span data-ttu-id="ec48f-919">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-919">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-920">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ec48f-920">Office 2016 on Windows</span></span><br><span data-ttu-id="ec48f-921">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-921">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-922">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-922">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-923">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-924">- Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-924">- Selection</span></span><br><span data-ttu-id="ec48f-925">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-925">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ec48f-926">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ec48f-926">Office 2013 on Windows</span></span><br><span data-ttu-id="ec48f-927">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ec48f-927">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ec48f-928">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ec48f-928">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ec48f-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ec48f-929">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ec48f-930">- Selection</span><span class="sxs-lookup"><span data-stu-id="ec48f-930">- Selection</span></span><br><span data-ttu-id="ec48f-931">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ec48f-931">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ec48f-932">関連項目</span><span class="sxs-lookup"><span data-stu-id="ec48f-932">See also</span></span>

- [<span data-ttu-id="ec48f-933">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="ec48f-933">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ec48f-934">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-934">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ec48f-935">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-935">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ec48f-936">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="ec48f-936">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ec48f-937">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="ec48f-937">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ec48f-938">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="ec48f-938">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ec48f-939">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="ec48f-939">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ec48f-940">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="ec48f-940">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ec48f-941">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="ec48f-941">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ec48f-942">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="ec48f-942">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ec48f-943">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="ec48f-943">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ec48f-944">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="ec48f-944">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)