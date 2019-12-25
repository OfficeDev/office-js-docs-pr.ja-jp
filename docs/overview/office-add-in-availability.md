---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 11/15/2019
localization_priority: Priority
ms.openlocfilehash: 956ee6b8a9e990a3d6d942ee4a65a1e9275ea025
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851370"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="7255b-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="7255b-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7255b-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="7255b-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7255b-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="7255b-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="7255b-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="7255b-108">Excel</span><span class="sxs-lookup"><span data-stu-id="7255b-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7255b-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7255b-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7255b-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7255b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="7255b-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-114">- TaskPane</span></span><br><span data-ttu-id="7255b-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-115">
        - Content</span></span><br><span data-ttu-id="7255b-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-116">
        - Custom Functions</span></span><br><span data-ttu-id="7255b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="7255b-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7255b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7255b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7255b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7255b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7255b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7255b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7255b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7255b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7255b-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7255b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="7255b-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="7255b-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="7255b-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="7255b-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-130">
        - BindingEvents</span></span><br><span data-ttu-id="7255b-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-131">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-132">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-133">
        - File</span></span><br><span data-ttu-id="7255b-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-134">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-136">
        - Selection</span></span><br><span data-ttu-id="7255b-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-137">
        - Settings</span></span><br><span data-ttu-id="7255b-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-138">
        - TableBindings</span></span><br><span data-ttu-id="7255b-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-139">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-140">
        - TextBindings</span></span><br><span data-ttu-id="7255b-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7255b-142">Office on Windows</span></span><br><span data-ttu-id="7255b-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-144">- TaskPane</span></span><br><span data-ttu-id="7255b-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-145">
        - Content</span></span><br><span data-ttu-id="7255b-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-146">
        - Custom Functions</span></span><br><span data-ttu-id="7255b-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="7255b-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="7255b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7255b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7255b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7255b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7255b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7255b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7255b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7255b-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7255b-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7255b-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="7255b-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="7255b-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7255b-161">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-161">
        - BindingEvents</span></span><br><span data-ttu-id="7255b-162">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-162">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-163">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-163">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-164">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-164">
        - File</span></span><br><span data-ttu-id="7255b-165">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-165">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-166">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-166">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-167">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-167">
        - Selection</span></span><br><span data-ttu-id="7255b-168">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-168">
        - Settings</span></span><br><span data-ttu-id="7255b-169">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-169">
        - TableBindings</span></span><br><span data-ttu-id="7255b-170">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-170">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-171">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-171">
        - TextBindings</span></span><br><span data-ttu-id="7255b-172">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-172">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-173">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-173">Office 2019 on Windows</span></span><br><span data-ttu-id="7255b-174">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-174">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7255b-175">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-175">- TaskPane</span></span><br><span data-ttu-id="7255b-176">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-176">
        - Content</span></span><br><span data-ttu-id="7255b-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7255b-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-178">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7255b-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7255b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7255b-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7255b-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7255b-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7255b-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7255b-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-188">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-188">- BindingEvents</span></span><br><span data-ttu-id="7255b-189">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-189">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-190">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-190">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-191">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-191">
        - File</span></span><br><span data-ttu-id="7255b-192">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-192">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-193">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-193">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-194">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-194">
        - Selection</span></span><br><span data-ttu-id="7255b-195">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-195">
        - Settings</span></span><br><span data-ttu-id="7255b-196">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-196">
        - TableBindings</span></span><br><span data-ttu-id="7255b-197">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-197">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-198">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-198">
        - TextBindings</span></span><br><span data-ttu-id="7255b-199">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-199">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-200">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-200">Office 2016 on Windows</span></span><br><span data-ttu-id="7255b-201">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-201">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7255b-202">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-202">- TaskPane</span></span><br><span data-ttu-id="7255b-203">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-203">
        - Content</span></span></td>
    <td><span data-ttu-id="7255b-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-204">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7255b-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-206">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-207">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-207">- BindingEvents</span></span><br><span data-ttu-id="7255b-208">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-208">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-209">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-209">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-210">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-210">
        - File</span></span><br><span data-ttu-id="7255b-211">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-211">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-212">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-212">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-213">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-213">
        - Selection</span></span><br><span data-ttu-id="7255b-214">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-214">
        - Settings</span></span><br><span data-ttu-id="7255b-215">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-215">
        - TableBindings</span></span><br><span data-ttu-id="7255b-216">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-216">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-217">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-217">
        - TextBindings</span></span><br><span data-ttu-id="7255b-218">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-218">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-219">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7255b-219">Office 2013 on Windows</span></span><br><span data-ttu-id="7255b-220">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-220">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7255b-221">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-221">
        - TaskPane</span></span><br><span data-ttu-id="7255b-222">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-222">
        - Content</span></span></td>
    <td>  <span data-ttu-id="7255b-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7255b-223">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7255b-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-224">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-225">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-225">
        - BindingEvents</span></span><br><span data-ttu-id="7255b-226">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-226">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-227">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-227">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-228">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-228">
        - File</span></span><br><span data-ttu-id="7255b-229">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-229">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-230">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-230">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-231">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-231">
        - Selection</span></span><br><span data-ttu-id="7255b-232">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-232">
        - Settings</span></span><br><span data-ttu-id="7255b-233">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-233">
        - TableBindings</span></span><br><span data-ttu-id="7255b-234">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-234">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-235">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-235">
        - TextBindings</span></span><br><span data-ttu-id="7255b-236">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-236">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-237">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-237">Office on iPad</span></span><br><span data-ttu-id="7255b-238">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-238">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7255b-239">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-239">- TaskPane</span></span><br><span data-ttu-id="7255b-240">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-240">
        - Content</span></span></td>
    <td><span data-ttu-id="7255b-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7255b-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7255b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7255b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7255b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7255b-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7255b-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7255b-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7255b-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7255b-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="7255b-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="7255b-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-253">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-253">- BindingEvents</span></span><br><span data-ttu-id="7255b-254">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-254">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-255">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-255">
        - File</span></span><br><span data-ttu-id="7255b-256">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-256">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-257">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-257">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-258">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-258">
        - Selection</span></span><br><span data-ttu-id="7255b-259">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-259">
        - Settings</span></span><br><span data-ttu-id="7255b-260">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-260">
        - TableBindings</span></span><br><span data-ttu-id="7255b-261">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-261">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-262">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-262">
        - TextBindings</span></span><br><span data-ttu-id="7255b-263">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-263">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-264">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-264">Office on Mac</span></span><br><span data-ttu-id="7255b-265">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-265">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7255b-266">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-266">- TaskPane</span></span><br><span data-ttu-id="7255b-267">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-267">
        - Content</span></span><br><span data-ttu-id="7255b-268">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-268">
        - Custom Functions</span></span><br><span data-ttu-id="7255b-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7255b-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-270">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7255b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7255b-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7255b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7255b-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7255b-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7255b-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7255b-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="7255b-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="7255b-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="7255b-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="7255b-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="7255b-283">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-283">- BindingEvents</span></span><br><span data-ttu-id="7255b-284">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-284">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-285">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-285">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-286">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-286">
        - File</span></span><br><span data-ttu-id="7255b-287">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-287">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-288">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-288">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-289">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-289">
        - PdfFile</span></span><br><span data-ttu-id="7255b-290">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-290">
        - Selection</span></span><br><span data-ttu-id="7255b-291">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-291">
        - Settings</span></span><br><span data-ttu-id="7255b-292">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-292">
        - TableBindings</span></span><br><span data-ttu-id="7255b-293">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-293">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-294">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-294">
        - TextBindings</span></span><br><span data-ttu-id="7255b-295">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-295">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-296">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-296">Office 2019 on Mac</span></span><br><span data-ttu-id="7255b-297">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-297">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7255b-298">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-298">- TaskPane</span></span><br><span data-ttu-id="7255b-299">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-299">
        - Content</span></span><br><span data-ttu-id="7255b-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="7255b-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-301">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="7255b-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="7255b-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="7255b-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="7255b-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="7255b-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="7255b-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="7255b-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-311">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-311">- BindingEvents</span></span><br><span data-ttu-id="7255b-312">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-312">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-313">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-313">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-314">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-314">
        - File</span></span><br><span data-ttu-id="7255b-315">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-315">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-316">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-316">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-317">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-317">
        - PdfFile</span></span><br><span data-ttu-id="7255b-318">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-318">
        - Selection</span></span><br><span data-ttu-id="7255b-319">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-319">
        - Settings</span></span><br><span data-ttu-id="7255b-320">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-320">
        - TableBindings</span></span><br><span data-ttu-id="7255b-321">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-321">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-322">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-322">
        - TextBindings</span></span><br><span data-ttu-id="7255b-323">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-323">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-324">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-324">Office 2016 on Mac</span></span><br><span data-ttu-id="7255b-325">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-325">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="7255b-326">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-326">- TaskPane</span></span><br><span data-ttu-id="7255b-327">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-327">
        - Content</span></span></td>
    <td><span data-ttu-id="7255b-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-328">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="7255b-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-329">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7255b-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-330">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="7255b-331">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-331">- BindingEvents</span></span><br><span data-ttu-id="7255b-332">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-332">
        - CompressedFile</span></span><br><span data-ttu-id="7255b-333">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-333">
        - DocumentEvents</span></span><br><span data-ttu-id="7255b-334">
        - File</span><span class="sxs-lookup"><span data-stu-id="7255b-334">
        - File</span></span><br><span data-ttu-id="7255b-335">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-335">
        - MatrixBindings</span></span><br><span data-ttu-id="7255b-336">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-336">
        - MatrixCoercion</span></span><br><span data-ttu-id="7255b-337">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-337">
        - PdfFile</span></span><br><span data-ttu-id="7255b-338">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-338">
        - Selection</span></span><br><span data-ttu-id="7255b-339">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-339">
        - Settings</span></span><br><span data-ttu-id="7255b-340">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-340">
        - TableBindings</span></span><br><span data-ttu-id="7255b-341">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-341">
        - TableCoercion</span></span><br><span data-ttu-id="7255b-342">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-342">
        - TextBindings</span></span><br><span data-ttu-id="7255b-343">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-343">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7255b-344">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7255b-344">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="7255b-345">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-345">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="7255b-346">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-346">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="7255b-347">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-347">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="7255b-348">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-348">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="7255b-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-349"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-350">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-350">Office on the web</span></span></td>
    <td><span data-ttu-id="7255b-351">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-351">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7255b-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-352">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-353">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7255b-353">Office on Windows</span></span><br><span data-ttu-id="7255b-354">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-354">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="7255b-355">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-355">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7255b-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-356">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-357">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="7255b-357">Office for Mac</span></span><br><span data-ttu-id="7255b-358">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="7255b-358">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="7255b-359">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="7255b-359">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="7255b-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-360">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="7255b-361">Outlook</span><span class="sxs-lookup"><span data-stu-id="7255b-361">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7255b-362">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-362">Platform</span></span></th>
    <th><span data-ttu-id="7255b-363">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-363">Extension points</span></span></th>
    <th><span data-ttu-id="7255b-364">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-364">API requirement sets</span></span></th>
    <th><span data-ttu-id="7255b-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-365"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-366">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-366">Office on the web</span></span><br><span data-ttu-id="7255b-367">(モダン)</span><span class="sxs-lookup"><span data-stu-id="7255b-367">(modern)</span></span></td>
    <td> <span data-ttu-id="7255b-368">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-368">- Mail Read</span></span><br><span data-ttu-id="7255b-369">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-369">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-370">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-371">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-376">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7255b-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="7255b-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="7255b-379">利用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-379">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-380">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-380">Office on the web</span></span><br><span data-ttu-id="7255b-381">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="7255b-381">(classic)</span></span></td>
    <td> <span data-ttu-id="7255b-382">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-382">- Mail Read</span></span><br><span data-ttu-id="7255b-383">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-383">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-384">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-385">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7255b-391">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-391">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-392">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7255b-392">Office on Windows</span></span><br><span data-ttu-id="7255b-393">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-393">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-394">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-394">- Mail Read</span></span><br><span data-ttu-id="7255b-395">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-395">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-396">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7255b-397">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="7255b-397">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7255b-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-398">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-402">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-403">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7255b-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-404">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="7255b-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="7255b-406">利用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-406">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-407">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-407">Office 2019 on Windows</span></span><br><span data-ttu-id="7255b-408">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-408">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-409">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-409">- Mail Read</span></span><br><span data-ttu-id="7255b-410">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-410">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-411">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7255b-412">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="7255b-412">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7255b-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-413">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-416">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-417">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-418">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7255b-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-419">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="7255b-420">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-420">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-421">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-421">Office 2016 on Windows</span></span><br><span data-ttu-id="7255b-422">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-422">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-423">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-423">- Mail Read</span></span><br><span data-ttu-id="7255b-424">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-424">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-425">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="7255b-426">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="7255b-426">
      - Modules</span></span></td>
    <td> <span data-ttu-id="7255b-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-427">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7255b-431">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-432">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7255b-432">Office 2013 on Windows</span></span><br><span data-ttu-id="7255b-433">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-434">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-434">- Mail Read</span></span><br><span data-ttu-id="7255b-435">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-435">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="7255b-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-436">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-437">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="7255b-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="7255b-440">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-440">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-441">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-441">Office on iOS</span></span><br><span data-ttu-id="7255b-442">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-442">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-443">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-443">- Mail Read</span></span><br><span data-ttu-id="7255b-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-444">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-445">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-446">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-447">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-448">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-449">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7255b-450">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-450">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-451">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-451">Office on Mac</span></span><br><span data-ttu-id="7255b-452">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-452">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-453">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-453">- Mail Read</span></span><br><span data-ttu-id="7255b-454">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-454">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-455">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-456">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-459">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-460">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="7255b-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="7255b-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="7255b-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="7255b-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="7255b-464">利用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-464">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-465">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-465">Office 2019 on Mac</span></span><br><span data-ttu-id="7255b-466">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-466">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-467">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-467">- Mail Read</span></span><br><span data-ttu-id="7255b-468">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-468">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-469">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-470">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-471">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-472">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-473">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7255b-476">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-476">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-477">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-477">Office 2016 on Mac</span></span><br><span data-ttu-id="7255b-478">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-478">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-479">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-479">- Mail Read</span></span><br><span data-ttu-id="7255b-480">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="7255b-480">
      - Mail Compose</span></span><br><span data-ttu-id="7255b-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-481">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-482">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-483">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-484">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-485">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="7255b-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="7255b-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="7255b-488">使用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-488">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-489">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-489">Office on Android</span></span><br><span data-ttu-id="7255b-490">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-490">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-491">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="7255b-491">- Mail Read</span></span><br><span data-ttu-id="7255b-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-492">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-493">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="7255b-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="7255b-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-495">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="7255b-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="7255b-496">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="7255b-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="7255b-497">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="7255b-498">利用不可</span><span class="sxs-lookup"><span data-stu-id="7255b-498">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="7255b-499">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7255b-499">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="7255b-500">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="7255b-500">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="7255b-501">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7255b-501">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="7255b-502">Word</span><span class="sxs-lookup"><span data-stu-id="7255b-502">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7255b-503">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-503">Platform</span></span></th>
    <th><span data-ttu-id="7255b-504">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-504">Extension points</span></span></th>
    <th><span data-ttu-id="7255b-505">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-505">API requirement sets</span></span></th>
    <th><span data-ttu-id="7255b-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-506"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-507">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-507">Office on the web</span></span></td>
    <td> <span data-ttu-id="7255b-508">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-508">- TaskPane</span></span><br><span data-ttu-id="7255b-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-509">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-510">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-511">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7255b-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-512">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7255b-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-513">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-514">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-515">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7255b-516">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-516">- BindingEvents</span></span><br><span data-ttu-id="7255b-517">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-517">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-518">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-518">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-519">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-519">
         - File</span></span><br><span data-ttu-id="7255b-520">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-520">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-521">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-521">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-522">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-522">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-523">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-523">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-524">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-524">
         - PdfFile</span></span><br><span data-ttu-id="7255b-525">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-525">
         - Selection</span></span><br><span data-ttu-id="7255b-526">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-526">
         - Settings</span></span><br><span data-ttu-id="7255b-527">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-527">
         - TableBindings</span></span><br><span data-ttu-id="7255b-528">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-528">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-529">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-529">
         - TextBindings</span></span><br><span data-ttu-id="7255b-530">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-530">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-531">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-531">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-532">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7255b-532">Office on Windows</span></span><br><span data-ttu-id="7255b-533">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-533">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-534">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-534">- TaskPane</span></span><br><span data-ttu-id="7255b-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-536">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-537">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7255b-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-538">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7255b-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-539">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-540">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-541">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7255b-542">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-542">- BindingEvents</span></span><br><span data-ttu-id="7255b-543">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-543">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-544">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-544">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-545">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-545">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-546">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-546">
         - File</span></span><br><span data-ttu-id="7255b-547">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-547">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-548">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-548">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-549">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-549">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-550">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-550">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-551">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-551">
         - PdfFile</span></span><br><span data-ttu-id="7255b-552">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-552">
         - Selection</span></span><br><span data-ttu-id="7255b-553">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-553">
         - Settings</span></span><br><span data-ttu-id="7255b-554">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-554">
         - TableBindings</span></span><br><span data-ttu-id="7255b-555">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-555">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-556">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-556">
         - TextBindings</span></span><br><span data-ttu-id="7255b-557">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-557">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-558">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-558">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-559">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-559">Office 2019 on Windows</span></span><br><span data-ttu-id="7255b-560">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-560">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-561">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-561">- TaskPane</span></span><br><span data-ttu-id="7255b-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-562">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-563">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-564">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7255b-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-565">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7255b-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-566">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-567">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-568">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-568">- BindingEvents</span></span><br><span data-ttu-id="7255b-569">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-569">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-570">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-570">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-571">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-571">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-572">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-572">
         - File</span></span><br><span data-ttu-id="7255b-573">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-573">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-574">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-574">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-575">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-575">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-576">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-576">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-577">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-577">
         - PdfFile</span></span><br><span data-ttu-id="7255b-578">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-578">
         - Selection</span></span><br><span data-ttu-id="7255b-579">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-579">
         - Settings</span></span><br><span data-ttu-id="7255b-580">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-580">
         - TableBindings</span></span><br><span data-ttu-id="7255b-581">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-581">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-582">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-582">
         - TextBindings</span></span><br><span data-ttu-id="7255b-583">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-583">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-584">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-584">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-585">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-585">Office 2016 on Windows</span></span><br><span data-ttu-id="7255b-586">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-586">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-587">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-587">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-588">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-589">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7255b-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-590">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-591">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-591">- BindingEvents</span></span><br><span data-ttu-id="7255b-592">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-592">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-593">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-593">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-594">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-594">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-595">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-595">
         - File</span></span><br><span data-ttu-id="7255b-596">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-596">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-597">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-597">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-598">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-598">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-599">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-599">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-600">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-600">
         - PdfFile</span></span><br><span data-ttu-id="7255b-601">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-601">
         - Selection</span></span><br><span data-ttu-id="7255b-602">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-602">
         - Settings</span></span><br><span data-ttu-id="7255b-603">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-603">
         - TableBindings</span></span><br><span data-ttu-id="7255b-604">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-604">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-605">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-605">
         - TextBindings</span></span><br><span data-ttu-id="7255b-606">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-606">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-607">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-607">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-608">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7255b-608">Office 2013 on Windows</span></span><br><span data-ttu-id="7255b-609">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-609">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-610">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-610">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7255b-611">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7255b-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-613">- BindingEvents</span></span><br><span data-ttu-id="7255b-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-614">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-616">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-617">
         - File</span></span><br><span data-ttu-id="7255b-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-619">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-622">
         - PdfFile</span></span><br><span data-ttu-id="7255b-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-623">
         - Selection</span></span><br><span data-ttu-id="7255b-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-624">
         - Settings</span></span><br><span data-ttu-id="7255b-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-625">
         - TableBindings</span></span><br><span data-ttu-id="7255b-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-626">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-627">
         - TextBindings</span></span><br><span data-ttu-id="7255b-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-628">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-629">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-630">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-630">Office on iPad</span></span><br><span data-ttu-id="7255b-631">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-631">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-632">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-633">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-634">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7255b-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-635">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7255b-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-636">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-637">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="7255b-638">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-638">- BindingEvents</span></span><br><span data-ttu-id="7255b-639">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-639">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-640">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-640">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-641">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-641">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-642">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-642">
         - File</span></span><br><span data-ttu-id="7255b-643">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-643">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-644">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-644">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-645">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-645">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-646">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-646">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-647">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-647">
         - PdfFile</span></span><br><span data-ttu-id="7255b-648">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-648">
         - Selection</span></span><br><span data-ttu-id="7255b-649">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-649">
         - Settings</span></span><br><span data-ttu-id="7255b-650">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-650">
         - TableBindings</span></span><br><span data-ttu-id="7255b-651">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-651">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-652">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-652">
         - TextBindings</span></span><br><span data-ttu-id="7255b-653">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-653">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-654">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-654">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-655">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-655">Office on Mac</span></span><br><span data-ttu-id="7255b-656">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-656">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-657">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-657">- TaskPane</span></span><br><span data-ttu-id="7255b-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-659">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-660">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7255b-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-661">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7255b-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-662">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-663">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-664">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="7255b-665">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-665">- BindingEvents</span></span><br><span data-ttu-id="7255b-666">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-666">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-667">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-667">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-668">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-668">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-669">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-669">
         - File</span></span><br><span data-ttu-id="7255b-670">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-670">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-671">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-671">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-672">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-672">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-673">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-673">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-674">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-674">
         - PdfFile</span></span><br><span data-ttu-id="7255b-675">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-675">
         - Selection</span></span><br><span data-ttu-id="7255b-676">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-676">
         - Settings</span></span><br><span data-ttu-id="7255b-677">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-677">
         - TableBindings</span></span><br><span data-ttu-id="7255b-678">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-678">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-679">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-679">
         - TextBindings</span></span><br><span data-ttu-id="7255b-680">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-680">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-681">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-681">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-682">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-682">Office 2019 on Mac</span></span><br><span data-ttu-id="7255b-683">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-683">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-684">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-684">- TaskPane</span></span><br><span data-ttu-id="7255b-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-685">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-686">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-687">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="7255b-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="7255b-688">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="7255b-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-689">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-690">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="7255b-691">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-691">- BindingEvents</span></span><br><span data-ttu-id="7255b-692">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-692">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-693">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-693">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-694">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-694">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-695">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-695">
         - File</span></span><br><span data-ttu-id="7255b-696">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-696">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-697">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-697">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-698">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-698">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-699">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-699">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-700">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-700">
         - PdfFile</span></span><br><span data-ttu-id="7255b-701">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-701">
         - Selection</span></span><br><span data-ttu-id="7255b-702">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-702">
         - Settings</span></span><br><span data-ttu-id="7255b-703">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-703">
         - TableBindings</span></span><br><span data-ttu-id="7255b-704">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-704">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-705">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-705">
         - TextBindings</span></span><br><span data-ttu-id="7255b-706">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-706">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-707">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-707">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-708">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-708">Office 2016 on Mac</span></span><br><span data-ttu-id="7255b-709">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-709">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-710">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-710">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-711">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="7255b-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="7255b-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="7255b-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-713">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-714">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-714">- BindingEvents</span></span><br><span data-ttu-id="7255b-715">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-715">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-716">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="7255b-716">
         - CustomXmlParts</span></span><br><span data-ttu-id="7255b-717">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-717">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-718">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-718">
         - File</span></span><br><span data-ttu-id="7255b-719">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-719">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-720">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-720">
         - MatrixBindings</span></span><br><span data-ttu-id="7255b-721">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-721">
         - MatrixCoercion</span></span><br><span data-ttu-id="7255b-722">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-722">
         - OoxmlCoercion</span></span><br><span data-ttu-id="7255b-723">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-723">
         - PdfFile</span></span><br><span data-ttu-id="7255b-724">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-724">
         - Selection</span></span><br><span data-ttu-id="7255b-725">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-725">
         - Settings</span></span><br><span data-ttu-id="7255b-726">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-726">
         - TableBindings</span></span><br><span data-ttu-id="7255b-727">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-727">
         - TableCoercion</span></span><br><span data-ttu-id="7255b-728">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="7255b-728">
         - TextBindings</span></span><br><span data-ttu-id="7255b-729">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-729">
         - TextCoercion</span></span><br><span data-ttu-id="7255b-730">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="7255b-730">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="7255b-731">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7255b-731">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="7255b-732">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="7255b-732">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7255b-733">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-733">Platform</span></span></th>
    <th><span data-ttu-id="7255b-734">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-734">Extension points</span></span></th>
    <th><span data-ttu-id="7255b-735">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-735">API requirement sets</span></span></th>
    <th><span data-ttu-id="7255b-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-736"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-737">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-737">Office on the web</span></span></td>
    <td> <span data-ttu-id="7255b-738">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-738">- Content</span></span><br><span data-ttu-id="7255b-739">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-739">
         - TaskPane</span></span><br><span data-ttu-id="7255b-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-740">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-741">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7255b-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-742">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-743">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-744">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7255b-745">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-745">- ActiveView</span></span><br><span data-ttu-id="7255b-746">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-746">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-747">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-747">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-748">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-748">
         - File</span></span><br><span data-ttu-id="7255b-749">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-749">
         - PdfFile</span></span><br><span data-ttu-id="7255b-750">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-750">
         - Selection</span></span><br><span data-ttu-id="7255b-751">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-751">
         - Settings</span></span><br><span data-ttu-id="7255b-752">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-752">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-753">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="7255b-753">Office on Windows</span></span><br><span data-ttu-id="7255b-754">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-754">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-755">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-755">- Content</span></span><br><span data-ttu-id="7255b-756">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-756">
         - TaskPane</span></span><br><span data-ttu-id="7255b-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-757">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-758">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7255b-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-759">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-760">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-761">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7255b-762">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-762">- ActiveView</span></span><br><span data-ttu-id="7255b-763">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-763">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-764">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-764">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-765">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-765">
         - File</span></span><br><span data-ttu-id="7255b-766">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-766">
         - PdfFile</span></span><br><span data-ttu-id="7255b-767">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-767">
         - Selection</span></span><br><span data-ttu-id="7255b-768">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-768">
         - Settings</span></span><br><span data-ttu-id="7255b-769">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-769">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-770">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-770">Office 2019 on Windows</span></span><br><span data-ttu-id="7255b-771">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-771">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-772">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-772">- Content</span></span><br><span data-ttu-id="7255b-773">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-773">
         - TaskPane</span></span><br><span data-ttu-id="7255b-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-774">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-775">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-776">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-777">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-777">- ActiveView</span></span><br><span data-ttu-id="7255b-778">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-778">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-779">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-779">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-780">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-780">
         - File</span></span><br><span data-ttu-id="7255b-781">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-781">
         - PdfFile</span></span><br><span data-ttu-id="7255b-782">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-782">
         - Selection</span></span><br><span data-ttu-id="7255b-783">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-783">
         - Settings</span></span><br><span data-ttu-id="7255b-784">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-784">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-785">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-785">Office 2016 on Windows</span></span><br><span data-ttu-id="7255b-786">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-786">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-787">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-787">- Content</span></span><br><span data-ttu-id="7255b-788">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-788">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7255b-789">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7255b-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-790">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-791">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-791">- ActiveView</span></span><br><span data-ttu-id="7255b-792">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-792">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-793">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-793">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-794">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-794">
         - File</span></span><br><span data-ttu-id="7255b-795">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-795">
         - PdfFile</span></span><br><span data-ttu-id="7255b-796">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-796">
         - Selection</span></span><br><span data-ttu-id="7255b-797">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-797">
         - Settings</span></span><br><span data-ttu-id="7255b-798">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-798">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-799">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7255b-799">Office 2013 on Windows</span></span><br><span data-ttu-id="7255b-800">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-800">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-801">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-801">- Content</span></span><br><span data-ttu-id="7255b-802">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-802">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="7255b-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7255b-803">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7255b-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-804">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-805">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-805">- ActiveView</span></span><br><span data-ttu-id="7255b-806">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-806">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-807">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-807">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-808">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-808">
         - File</span></span><br><span data-ttu-id="7255b-809">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-809">
         - PdfFile</span></span><br><span data-ttu-id="7255b-810">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-810">
         - Selection</span></span><br><span data-ttu-id="7255b-811">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-811">
         - Settings</span></span><br><span data-ttu-id="7255b-812">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-812">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-813">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-813">Office on iPad</span></span><br><span data-ttu-id="7255b-814">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-814">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-815">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-815">- Content</span></span><br><span data-ttu-id="7255b-816">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-816">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-817">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7255b-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-818">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-819">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-820">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-820">- ActiveView</span></span><br><span data-ttu-id="7255b-821">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-821">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-822">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-822">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-823">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-823">
         - File</span></span><br><span data-ttu-id="7255b-824">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-824">
         - PdfFile</span></span><br><span data-ttu-id="7255b-825">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-825">
         - Selection</span></span><br><span data-ttu-id="7255b-826">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-826">
         - Settings</span></span><br><span data-ttu-id="7255b-827">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-827">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-828">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="7255b-828">Office on Mac</span></span><br><span data-ttu-id="7255b-829">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="7255b-829">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="7255b-830">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-830">- Content</span></span><br><span data-ttu-id="7255b-831">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-831">
         - TaskPane</span></span><br><span data-ttu-id="7255b-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-832">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-833">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="7255b-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-834">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-835">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="7255b-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="7255b-836">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="7255b-837">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-837">- ActiveView</span></span><br><span data-ttu-id="7255b-838">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-838">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-839">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-839">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-840">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-840">
         - File</span></span><br><span data-ttu-id="7255b-841">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-841">
         - PdfFile</span></span><br><span data-ttu-id="7255b-842">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-842">
         - Selection</span></span><br><span data-ttu-id="7255b-843">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-843">
         - Settings</span></span><br><span data-ttu-id="7255b-844">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-844">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-845">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-845">Office 2019 on Mac</span></span><br><span data-ttu-id="7255b-846">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-846">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-847">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-847">- Content</span></span><br><span data-ttu-id="7255b-848">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-848">
         - TaskPane</span></span><br><span data-ttu-id="7255b-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-849">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-850">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-851">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-852">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-852">- ActiveView</span></span><br><span data-ttu-id="7255b-853">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-853">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-854">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-854">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-855">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-855">
         - File</span></span><br><span data-ttu-id="7255b-856">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-856">
         - PdfFile</span></span><br><span data-ttu-id="7255b-857">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-857">
         - Selection</span></span><br><span data-ttu-id="7255b-858">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-858">
         - Settings</span></span><br><span data-ttu-id="7255b-859">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-859">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-860">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-860">Office 2016 on Mac</span></span><br><span data-ttu-id="7255b-861">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-861">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-862">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-862">- Content</span></span><br><span data-ttu-id="7255b-863">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-863">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="7255b-864">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="7255b-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-865">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-866">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="7255b-866">- ActiveView</span></span><br><span data-ttu-id="7255b-867">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="7255b-867">
         - CompressedFile</span></span><br><span data-ttu-id="7255b-868">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-868">
         - DocumentEvents</span></span><br><span data-ttu-id="7255b-869">
         - File</span><span class="sxs-lookup"><span data-stu-id="7255b-869">
         - File</span></span><br><span data-ttu-id="7255b-870">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="7255b-870">
         - PdfFile</span></span><br><span data-ttu-id="7255b-871">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-871">
         - Selection</span></span><br><span data-ttu-id="7255b-872">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-872">
         - Settings</span></span><br><span data-ttu-id="7255b-873">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-873">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="7255b-874">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="7255b-874">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="7255b-875">OneNote</span><span class="sxs-lookup"><span data-stu-id="7255b-875">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7255b-876">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-876">Platform</span></span></th>
    <th><span data-ttu-id="7255b-877">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-877">Extension points</span></span></th>
    <th><span data-ttu-id="7255b-878">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-878">API requirement sets</span></span></th>
    <th><span data-ttu-id="7255b-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-879"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-880">Office on the web</span><span class="sxs-lookup"><span data-stu-id="7255b-880">Office on the web</span></span></td>
    <td> <span data-ttu-id="7255b-881">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="7255b-881">- Content</span></span><br><span data-ttu-id="7255b-882">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-882">
         - TaskPane</span></span><br><span data-ttu-id="7255b-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="7255b-883">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="7255b-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-884">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="7255b-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-885">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="7255b-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-886">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-887">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="7255b-887">- DocumentEvents</span></span><br><span data-ttu-id="7255b-888">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-888">
         - HtmlCoercion</span></span><br><span data-ttu-id="7255b-889">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="7255b-889">
         - Settings</span></span><br><span data-ttu-id="7255b-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-890">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="7255b-891">Project</span><span class="sxs-lookup"><span data-stu-id="7255b-891">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="7255b-892">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="7255b-892">Platform</span></span></th>
    <th><span data-ttu-id="7255b-893">拡張点</span><span class="sxs-lookup"><span data-stu-id="7255b-893">Extension points</span></span></th>
    <th><span data-ttu-id="7255b-894">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-894">API requirement sets</span></span></th>
    <th><span data-ttu-id="7255b-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="7255b-895"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-896">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="7255b-896">Office 2019 on Windows</span></span><br><span data-ttu-id="7255b-897">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-897">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-898">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-898">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-899">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-900">- Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-900">- Selection</span></span><br><span data-ttu-id="7255b-901">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-901">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-902">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="7255b-902">Office 2016 on Windows</span></span><br><span data-ttu-id="7255b-903">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-903">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-904">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-904">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-905">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-906">- Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-906">- Selection</span></span><br><span data-ttu-id="7255b-907">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-907">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="7255b-908">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="7255b-908">Office 2013 on Windows</span></span><br><span data-ttu-id="7255b-909">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="7255b-909">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="7255b-910">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7255b-910">- TaskPane</span></span></td>
    <td> <span data-ttu-id="7255b-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="7255b-911">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="7255b-912">- Selection</span><span class="sxs-lookup"><span data-stu-id="7255b-912">- Selection</span></span><br><span data-ttu-id="7255b-913">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="7255b-913">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="7255b-914">関連項目</span><span class="sxs-lookup"><span data-stu-id="7255b-914">See also</span></span>

- [<span data-ttu-id="7255b-915">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="7255b-915">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="7255b-916">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-916">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="7255b-917">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-917">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="7255b-918">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="7255b-918">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="7255b-919">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="7255b-919">API reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="7255b-920">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="7255b-920">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="7255b-921">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="7255b-921">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="7255b-922">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="7255b-922">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="7255b-923">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="7255b-923">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="7255b-924">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="7255b-924">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="7255b-925">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="7255b-925">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="7255b-926">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="7255b-926">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)