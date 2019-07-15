---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 07/11/2019
localization_priority: Priority
ms.openlocfilehash: d88f7c1b9daa201d9b6bc5cfa69ac3125bf127b1
ms.sourcegitcommit: 61f8f02193ce05da957418d938f0d94cb12c468d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2019
ms.locfileid: "35630537"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="a5231-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="a5231-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a5231-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="a5231-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a5231-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="a5231-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="a5231-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="a5231-108">Excel</span><span class="sxs-lookup"><span data-stu-id="a5231-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="a5231-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="a5231-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="a5231-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="a5231-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="a5231-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-114">- TaskPane</span></span><br><span data-ttu-id="a5231-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-115">
        - Content</span></span><br><span data-ttu-id="a5231-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-116">
        - Custom Functions</span></span><br><span data-ttu-id="a5231-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="a5231-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a5231-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a5231-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a5231-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a5231-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a5231-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a5231-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a5231-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a5231-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a5231-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a5231-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a5231-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="a5231-130">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-130">
        - BindingEvents</span></span><br><span data-ttu-id="a5231-131">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-131">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-132">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-132">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-133">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-133">
        - File</span></span><br><span data-ttu-id="a5231-134">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-134">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-135">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-135">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-136">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-136">
        - Selection</span></span><br><span data-ttu-id="a5231-137">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-137">
        - Settings</span></span><br><span data-ttu-id="a5231-138">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-138">
        - TableBindings</span></span><br><span data-ttu-id="a5231-139">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-139">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-140">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-140">
        - TextBindings</span></span><br><span data-ttu-id="a5231-141">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-141">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-142">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="a5231-142">Office on Windows</span></span><br><span data-ttu-id="a5231-143">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-143">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-144">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-144">- TaskPane</span></span><br><span data-ttu-id="a5231-145">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-145">
        - Content</span></span><br><span data-ttu-id="a5231-146">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-146">
        - Custom Functions</span></span><br><span data-ttu-id="a5231-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="a5231-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="a5231-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a5231-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a5231-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a5231-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a5231-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a5231-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a5231-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a5231-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a5231-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a5231-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a5231-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="a5231-160">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-160">
        - BindingEvents</span></span><br><span data-ttu-id="a5231-161">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-161">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-162">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-162">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-163">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-163">
        - File</span></span><br><span data-ttu-id="a5231-164">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-164">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-165">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-165">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-166">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-166">
        - Selection</span></span><br><span data-ttu-id="a5231-167">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-167">
        - Settings</span></span><br><span data-ttu-id="a5231-168">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-168">
        - TableBindings</span></span><br><span data-ttu-id="a5231-169">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-169">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-170">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-170">
        - TextBindings</span></span><br><span data-ttu-id="a5231-171">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-171">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-172">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-172">Office 2019 on Windows</span></span><br><span data-ttu-id="a5231-173">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-173">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a5231-174">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-174">- TaskPane</span></span><br><span data-ttu-id="a5231-175">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-175">
        - Content</span></span><br><span data-ttu-id="a5231-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a5231-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-177">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a5231-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a5231-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a5231-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a5231-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a5231-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a5231-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a5231-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a5231-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="a5231-187">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-187">- BindingEvents</span></span><br><span data-ttu-id="a5231-188">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-188">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-189">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-189">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-190">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-190">
        - File</span></span><br><span data-ttu-id="a5231-191">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-191">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-192">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-192">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-193">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-193">
        - Selection</span></span><br><span data-ttu-id="a5231-194">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-194">
        - Settings</span></span><br><span data-ttu-id="a5231-195">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-195">
        - TableBindings</span></span><br><span data-ttu-id="a5231-196">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-196">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-197">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-197">
        - TextBindings</span></span><br><span data-ttu-id="a5231-198">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-198">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-199">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-199">Office 2016 on Windows</span></span><br><span data-ttu-id="a5231-200">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-200">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a5231-201">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-201">- TaskPane</span></span><br><span data-ttu-id="a5231-202">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-202">
        - Content</span></span></td>
    <td><span data-ttu-id="a5231-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-203">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-204">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="a5231-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-205">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="a5231-206">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-206">- BindingEvents</span></span><br><span data-ttu-id="a5231-207">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-207">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-208">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-208">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-209">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-209">
        - File</span></span><br><span data-ttu-id="a5231-210">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-210">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-211">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-211">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-212">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-212">
        - Selection</span></span><br><span data-ttu-id="a5231-213">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-213">
        - Settings</span></span><br><span data-ttu-id="a5231-214">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-214">
        - TableBindings</span></span><br><span data-ttu-id="a5231-215">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-215">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-216">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-216">
        - TextBindings</span></span><br><span data-ttu-id="a5231-217">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-217">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-218">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a5231-218">Office 2013 on Windows</span></span><br><span data-ttu-id="a5231-219">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-219">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a5231-220">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-220">
        - TaskPane</span></span><br><span data-ttu-id="a5231-221">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-221">
        - Content</span></span></td>
    <td>  <span data-ttu-id="a5231-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a5231-222">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="a5231-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-223">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="a5231-224">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-224">
        - BindingEvents</span></span><br><span data-ttu-id="a5231-225">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-225">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-226">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-226">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-227">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-227">
        - File</span></span><br><span data-ttu-id="a5231-228">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-228">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-229">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-229">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-230">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-230">
        - Selection</span></span><br><span data-ttu-id="a5231-231">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-231">
        - Settings</span></span><br><span data-ttu-id="a5231-232">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-232">
        - TableBindings</span></span><br><span data-ttu-id="a5231-233">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-233">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-234">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-234">
        - TextBindings</span></span><br><span data-ttu-id="a5231-235">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-235">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-236">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-236">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="a5231-237">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-237">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="a5231-238">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-238">- TaskPane</span></span><br><span data-ttu-id="a5231-239">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-239">
        - Content</span></span><br><span data-ttu-id="a5231-240">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-240">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="a5231-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-241">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a5231-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a5231-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a5231-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a5231-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a5231-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a5231-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a5231-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a5231-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a5231-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a5231-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="a5231-252">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-252">- BindingEvents</span></span><br><span data-ttu-id="a5231-253">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-253">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-254">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-254">
        - File</span></span><br><span data-ttu-id="a5231-255">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-255">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-256">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-256">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-257">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-257">
        - Selection</span></span><br><span data-ttu-id="a5231-258">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-258">
        - Settings</span></span><br><span data-ttu-id="a5231-259">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-259">
        - TableBindings</span></span><br><span data-ttu-id="a5231-260">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-260">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-261">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-261">
        - TextBindings</span></span><br><span data-ttu-id="a5231-262">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-262">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-263">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-263">Office apps on Mac</span></span><br><span data-ttu-id="a5231-264">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-264">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="a5231-265">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-265">- TaskPane</span></span><br><span data-ttu-id="a5231-266">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-266">
        - Content</span></span><br><span data-ttu-id="a5231-267">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-267">
        - Custom Functions</span></span><br><span data-ttu-id="a5231-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a5231-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-269">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a5231-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a5231-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a5231-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a5231-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a5231-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a5231-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a5231-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a5231-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="a5231-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.9</a></span></span><br><span data-ttu-id="a5231-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="a5231-281">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-281">- BindingEvents</span></span><br><span data-ttu-id="a5231-282">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-282">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-283">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-283">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-284">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-284">
        - File</span></span><br><span data-ttu-id="a5231-285">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-285">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-286">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-286">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-287">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-287">
        - PdfFile</span></span><br><span data-ttu-id="a5231-288">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-288">
        - Selection</span></span><br><span data-ttu-id="a5231-289">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-289">
        - Settings</span></span><br><span data-ttu-id="a5231-290">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-290">
        - TableBindings</span></span><br><span data-ttu-id="a5231-291">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-291">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-292">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-292">
        - TextBindings</span></span><br><span data-ttu-id="a5231-293">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-293">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-294">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-294">Office 2019 for Mac</span></span><br><span data-ttu-id="a5231-295">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-295">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a5231-296">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-296">- TaskPane</span></span><br><span data-ttu-id="a5231-297">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-297">
        - Content</span></span><br><span data-ttu-id="a5231-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="a5231-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-299">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a></span></span><br><span data-ttu-id="a5231-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a></span></span><br><span data-ttu-id="a5231-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a></span></span><br><span data-ttu-id="a5231-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.5</a></span></span><br><span data-ttu-id="a5231-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.6</a></span></span><br><span data-ttu-id="a5231-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.7</a></span></span><br><span data-ttu-id="a5231-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="a5231-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.8</a></span></span><br><span data-ttu-id="a5231-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="a5231-309">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-309">- BindingEvents</span></span><br><span data-ttu-id="a5231-310">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-310">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-311">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-311">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-312">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-312">
        - File</span></span><br><span data-ttu-id="a5231-313">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-313">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-314">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-314">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-315">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-315">
        - PdfFile</span></span><br><span data-ttu-id="a5231-316">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-316">
        - Selection</span></span><br><span data-ttu-id="a5231-317">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-317">
        - Settings</span></span><br><span data-ttu-id="a5231-318">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-318">
        - TableBindings</span></span><br><span data-ttu-id="a5231-319">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-319">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-320">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-320">
        - TextBindings</span></span><br><span data-ttu-id="a5231-321">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-321">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-322">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-322">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="a5231-323">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-323">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="a5231-324">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-324">- TaskPane</span></span><br><span data-ttu-id="a5231-325">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-325">
        - Content</span></span></td>
    <td><span data-ttu-id="a5231-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-326">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a></span></span><br><span data-ttu-id="a5231-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-327">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="a5231-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-328">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="a5231-329">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-329">- BindingEvents</span></span><br><span data-ttu-id="a5231-330">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-330">
        - CompressedFile</span></span><br><span data-ttu-id="a5231-331">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-331">
        - DocumentEvents</span></span><br><span data-ttu-id="a5231-332">
        - File</span><span class="sxs-lookup"><span data-stu-id="a5231-332">
        - File</span></span><br><span data-ttu-id="a5231-333">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-333">
        - MatrixBindings</span></span><br><span data-ttu-id="a5231-334">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-334">
        - MatrixCoercion</span></span><br><span data-ttu-id="a5231-335">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-335">
        - PdfFile</span></span><br><span data-ttu-id="a5231-336">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-336">
        - Selection</span></span><br><span data-ttu-id="a5231-337">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-337">
        - Settings</span></span><br><span data-ttu-id="a5231-338">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-338">
        - TableBindings</span></span><br><span data-ttu-id="a5231-339">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-339">
        - TableCoercion</span></span><br><span data-ttu-id="a5231-340">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-340">
        - TextBindings</span></span><br><span data-ttu-id="a5231-341">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-341">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="a5231-342">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a5231-342">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="a5231-343">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-343">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="a5231-344">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-344">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="a5231-345">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-345">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="a5231-346">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-346">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="a5231-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-347"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-348">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-348">Office on the web</span></span></td>
    <td><span data-ttu-id="a5231-349">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-349">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="a5231-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-350">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-351">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="a5231-351">Office on Windows</span></span><br><span data-ttu-id="a5231-352">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-352">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="a5231-353">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-353">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="a5231-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-354">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-355">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a5231-355">Office for Mac</span></span><br><span data-ttu-id="a5231-356">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="a5231-356">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="a5231-357">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="a5231-357">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="a5231-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-358">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="a5231-359">Outlook</span><span class="sxs-lookup"><span data-stu-id="a5231-359">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a5231-360">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-360">Platform</span></span></th>
    <th><span data-ttu-id="a5231-361">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-361">Extension points</span></span></th>
    <th><span data-ttu-id="a5231-362">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-362">API requirement sets</span></span></th>
    <th><span data-ttu-id="a5231-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-363"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-364">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-364">Office on the web</span></span><br><span data-ttu-id="a5231-365">(新規)</span><span class="sxs-lookup"><span data-stu-id="a5231-365">New</span></span></td>
    <td> <span data-ttu-id="a5231-366">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-366">- Mail Read</span></span><br><span data-ttu-id="a5231-367">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-367">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-368">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-369">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-373">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-374">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a5231-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-375">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a5231-376">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-376">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-377">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-377">Office on the web</span></span><br><span data-ttu-id="a5231-378">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="a5231-378">Classic</span></span></td>
    <td> <span data-ttu-id="a5231-379">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-379">- Mail Read</span></span><br><span data-ttu-id="a5231-380">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-380">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-381">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-382">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-385">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-386">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-387">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a5231-388">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-388">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-389">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="a5231-389">Office on Windows</span></span><br><span data-ttu-id="a5231-390">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-390">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-391">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-391">- Mail Read</span></span><br><span data-ttu-id="a5231-392">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-392">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-393">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a5231-394">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="a5231-394">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a5231-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-395">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-399">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-400">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a5231-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-401">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a5231-402">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-402">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-403">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-403">Office 2019 on Windows</span></span><br><span data-ttu-id="a5231-404">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-404">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-405">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-405">- Mail Read</span></span><br><span data-ttu-id="a5231-406">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-406">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-407">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a5231-408">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="a5231-408">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a5231-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-409">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a5231-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-415">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a5231-416">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-416">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-417">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-417">Office 2016 on Windows</span></span><br><span data-ttu-id="a5231-418">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-418">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-419">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-419">- Mail Read</span></span><br><span data-ttu-id="a5231-420">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-420">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-421">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="a5231-422">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="a5231-422">
      - Modules</span></span></td>
    <td> <span data-ttu-id="a5231-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-423">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="a5231-427">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-427">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-428">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a5231-428">Office 2013 on Windows</span></span><br><span data-ttu-id="a5231-429">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-429">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-430">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-430">- Mail Read</span></span><br><span data-ttu-id="a5231-431">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-431">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="a5231-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-432">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-433">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-434">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="a5231-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-435">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="a5231-436">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-436">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-437">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-437">Office apps on iOS</span></span><br><span data-ttu-id="a5231-438">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-438">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-439">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-439">- Mail Read</span></span><br><span data-ttu-id="a5231-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-440">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-441">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-445">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a5231-446">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-446">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-447">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-447">Office apps on Mac</span></span><br><span data-ttu-id="a5231-448">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-448">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-449">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-449">- Mail Read</span></span><br><span data-ttu-id="a5231-450">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-450">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-451">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-452">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="a5231-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="a5231-458">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="a5231-459">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-459">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-460">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-460">Office 2019 for Mac</span></span><br><span data-ttu-id="a5231-461">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-461">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-462">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-462">- Mail Read</span></span><br><span data-ttu-id="a5231-463">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-463">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-464">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-465">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-469">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-470">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a5231-471">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-471">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-472">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-472">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="a5231-473">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-473">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-474">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-474">- Mail Read</span></span><br><span data-ttu-id="a5231-475">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="a5231-475">
      - Mail Compose</span></span><br><span data-ttu-id="a5231-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-476">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-477">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-481">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="a5231-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="a5231-482">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="a5231-483">使用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-483">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-484">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-484">Office apps on Android</span></span><br><span data-ttu-id="a5231-485">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-485">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-486">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="a5231-486">- Mail Read</span></span><br><span data-ttu-id="a5231-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-487">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-488">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="a5231-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="a5231-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="a5231-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="a5231-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="a5231-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="a5231-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="a5231-493">利用不可</span><span class="sxs-lookup"><span data-stu-id="a5231-493">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="a5231-494">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a5231-494">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="a5231-495">Word</span><span class="sxs-lookup"><span data-stu-id="a5231-495">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a5231-496">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-496">Platform</span></span></th>
    <th><span data-ttu-id="a5231-497">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-497">Extension points</span></span></th>
    <th><span data-ttu-id="a5231-498">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-498">API requirement sets</span></span></th>
    <th><span data-ttu-id="a5231-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-499"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-500">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-500">Office on the web</span></span></td>
    <td> <span data-ttu-id="a5231-501">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-501">- TaskPane</span></span><br><span data-ttu-id="a5231-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-502">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-503">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a5231-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a5231-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-508">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="a5231-509">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-509">- BindingEvents</span></span><br><span data-ttu-id="a5231-510">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-510">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-511">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-511">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-512">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-512">
         - File</span></span><br><span data-ttu-id="a5231-513">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-513">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-514">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-514">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-515">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-515">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-516">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-516">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-517">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-517">
         - PdfFile</span></span><br><span data-ttu-id="a5231-518">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-518">
         - Selection</span></span><br><span data-ttu-id="a5231-519">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-519">
         - Settings</span></span><br><span data-ttu-id="a5231-520">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-520">
         - TableBindings</span></span><br><span data-ttu-id="a5231-521">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-521">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-522">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-522">
         - TextBindings</span></span><br><span data-ttu-id="a5231-523">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-523">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-524">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-524">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-525">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="a5231-525">Office on Windows</span></span><br><span data-ttu-id="a5231-526">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-526">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-527">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-527">- TaskPane</span></span><br><span data-ttu-id="a5231-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-528">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-529">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a5231-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a5231-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-532">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-534">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="a5231-535">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-535">- BindingEvents</span></span><br><span data-ttu-id="a5231-536">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-536">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-537">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-537">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-538">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-538">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-539">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-539">
         - File</span></span><br><span data-ttu-id="a5231-540">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-540">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-541">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-541">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-542">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-542">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-543">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-543">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-544">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-544">
         - PdfFile</span></span><br><span data-ttu-id="a5231-545">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-545">
         - Selection</span></span><br><span data-ttu-id="a5231-546">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-546">
         - Settings</span></span><br><span data-ttu-id="a5231-547">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-547">
         - TableBindings</span></span><br><span data-ttu-id="a5231-548">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-548">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-549">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-549">
         - TextBindings</span></span><br><span data-ttu-id="a5231-550">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-550">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-551">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-551">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-552">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-552">Office 2019 on Windows</span></span><br><span data-ttu-id="a5231-553">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-553">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-554">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-554">- TaskPane</span></span><br><span data-ttu-id="a5231-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-555">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-556">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a5231-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a5231-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-560">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-561">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-561">- BindingEvents</span></span><br><span data-ttu-id="a5231-562">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-562">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-563">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-563">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-564">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-564">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-565">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-565">
         - File</span></span><br><span data-ttu-id="a5231-566">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-566">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-567">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-567">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-568">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-568">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-569">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-569">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-570">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-570">
         - PdfFile</span></span><br><span data-ttu-id="a5231-571">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-571">
         - Selection</span></span><br><span data-ttu-id="a5231-572">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-572">
         - Settings</span></span><br><span data-ttu-id="a5231-573">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-573">
         - TableBindings</span></span><br><span data-ttu-id="a5231-574">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-574">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-575">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-575">
         - TextBindings</span></span><br><span data-ttu-id="a5231-576">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-576">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-577">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-577">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-578">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-578">Office 2016 on Windows</span></span><br><span data-ttu-id="a5231-579">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-579">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-580">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-580">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-581">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-582">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="a5231-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-583">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-584">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-584">- BindingEvents</span></span><br><span data-ttu-id="a5231-585">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-585">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-586">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-586">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-587">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-587">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-588">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-588">
         - File</span></span><br><span data-ttu-id="a5231-589">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-589">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-590">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-590">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-591">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-591">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-592">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-592">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-593">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-593">
         - PdfFile</span></span><br><span data-ttu-id="a5231-594">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-594">
         - Selection</span></span><br><span data-ttu-id="a5231-595">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-595">
         - Settings</span></span><br><span data-ttu-id="a5231-596">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-596">
         - TableBindings</span></span><br><span data-ttu-id="a5231-597">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-597">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-598">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-598">
         - TextBindings</span></span><br><span data-ttu-id="a5231-599">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-599">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-600">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-600">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-601">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a5231-601">Office 2013 on Windows</span></span><br><span data-ttu-id="a5231-602">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-602">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-603">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-603">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a5231-604">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="a5231-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-605">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-606">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-606">- BindingEvents</span></span><br><span data-ttu-id="a5231-607">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-607">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-608">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-608">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-609">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-609">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-610">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-610">
         - File</span></span><br><span data-ttu-id="a5231-611">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-611">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-612">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-612">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-613">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-613">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-614">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-614">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-615">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-615">
         - PdfFile</span></span><br><span data-ttu-id="a5231-616">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-616">
         - Selection</span></span><br><span data-ttu-id="a5231-617">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-617">
         - Settings</span></span><br><span data-ttu-id="a5231-618">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-618">
         - TableBindings</span></span><br><span data-ttu-id="a5231-619">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-619">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-620">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-620">
         - TextBindings</span></span><br><span data-ttu-id="a5231-621">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-621">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-622">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-622">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-623">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-623">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="a5231-624">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-624">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-625">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-625">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-626">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a5231-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a5231-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-630">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="a5231-631">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-631">- BindingEvents</span></span><br><span data-ttu-id="a5231-632">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-632">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-633">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-633">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-634">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-634">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-635">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-635">
         - File</span></span><br><span data-ttu-id="a5231-636">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-636">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-637">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-637">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-638">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-638">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-639">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-639">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-640">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-640">
         - PdfFile</span></span><br><span data-ttu-id="a5231-641">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-641">
         - Selection</span></span><br><span data-ttu-id="a5231-642">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-642">
         - Settings</span></span><br><span data-ttu-id="a5231-643">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-643">
         - TableBindings</span></span><br><span data-ttu-id="a5231-644">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-644">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-645">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-645">
         - TextBindings</span></span><br><span data-ttu-id="a5231-646">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-646">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-647">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-647">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-648">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-648">Office apps on Mac</span></span><br><span data-ttu-id="a5231-649">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-649">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-650">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-650">- TaskPane</span></span><br><span data-ttu-id="a5231-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-651">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-652">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a5231-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a5231-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-657">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="a5231-658">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-658">- BindingEvents</span></span><br><span data-ttu-id="a5231-659">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-659">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-660">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-660">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-661">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-661">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-662">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-662">
         - File</span></span><br><span data-ttu-id="a5231-663">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-663">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-664">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-664">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-665">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-665">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-666">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-666">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-667">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-667">
         - PdfFile</span></span><br><span data-ttu-id="a5231-668">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-668">
         - Selection</span></span><br><span data-ttu-id="a5231-669">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-669">
         - Settings</span></span><br><span data-ttu-id="a5231-670">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-670">
         - TableBindings</span></span><br><span data-ttu-id="a5231-671">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-671">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-672">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-672">
         - TextBindings</span></span><br><span data-ttu-id="a5231-673">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-673">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-674">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-674">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-675">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-675">Office 2019 for Mac</span></span><br><span data-ttu-id="a5231-676">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-676">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-677">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-677">- TaskPane</span></span><br><span data-ttu-id="a5231-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-678">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-679">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.2</a></span></span><br><span data-ttu-id="a5231-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="a5231-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.3</a></span></span><br><span data-ttu-id="a5231-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="a5231-684">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-684">- BindingEvents</span></span><br><span data-ttu-id="a5231-685">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-685">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-686">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-686">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-687">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-687">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-688">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-688">
         - File</span></span><br><span data-ttu-id="a5231-689">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-689">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-690">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-690">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-691">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-691">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-692">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-692">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-693">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-693">
         - PdfFile</span></span><br><span data-ttu-id="a5231-694">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-694">
         - Selection</span></span><br><span data-ttu-id="a5231-695">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-695">
         - Settings</span></span><br><span data-ttu-id="a5231-696">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-696">
         - TableBindings</span></span><br><span data-ttu-id="a5231-697">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-697">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-698">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-698">
         - TextBindings</span></span><br><span data-ttu-id="a5231-699">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-699">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-700">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-700">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-701">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-701">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="a5231-702">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-702">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-703">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-703">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-704">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets">WordApi 1.1</a></span></span><br><span data-ttu-id="a5231-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="a5231-705">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="a5231-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-706">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-707">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-707">- BindingEvents</span></span><br><span data-ttu-id="a5231-708">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-708">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-709">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="a5231-709">
         - CustomXmlParts</span></span><br><span data-ttu-id="a5231-710">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-710">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-711">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-711">
         - File</span></span><br><span data-ttu-id="a5231-712">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-712">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-713">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-713">
         - MatrixBindings</span></span><br><span data-ttu-id="a5231-714">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-714">
         - MatrixCoercion</span></span><br><span data-ttu-id="a5231-715">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-715">
         - OoxmlCoercion</span></span><br><span data-ttu-id="a5231-716">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-716">
         - PdfFile</span></span><br><span data-ttu-id="a5231-717">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-717">
         - Selection</span></span><br><span data-ttu-id="a5231-718">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-718">
         - Settings</span></span><br><span data-ttu-id="a5231-719">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-719">
         - TableBindings</span></span><br><span data-ttu-id="a5231-720">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-720">
         - TableCoercion</span></span><br><span data-ttu-id="a5231-721">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="a5231-721">
         - TextBindings</span></span><br><span data-ttu-id="a5231-722">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-722">
         - TextCoercion</span></span><br><span data-ttu-id="a5231-723">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="a5231-723">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="a5231-724">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a5231-724">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="a5231-725">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="a5231-725">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a5231-726">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-726">Platform</span></span></th>
    <th><span data-ttu-id="a5231-727">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-727">Extension points</span></span></th>
    <th><span data-ttu-id="a5231-728">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-728">API requirement sets</span></span></th>
    <th><span data-ttu-id="a5231-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-729"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-730">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-730">Office on the web</span></span></td>
    <td> <span data-ttu-id="a5231-731">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-731">- Content</span></span><br><span data-ttu-id="a5231-732">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-732">
         - TaskPane</span></span><br><span data-ttu-id="a5231-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-733">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-734">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-736">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="a5231-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-737">- ActiveView</span></span><br><span data-ttu-id="a5231-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-738">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-739">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-740">
         - File</span></span><br><span data-ttu-id="a5231-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-741">
         - PdfFile</span></span><br><span data-ttu-id="a5231-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-742">
         - Selection</span></span><br><span data-ttu-id="a5231-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-743">
         - Settings</span></span><br><span data-ttu-id="a5231-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-745">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="a5231-745">Office on Windows</span></span><br><span data-ttu-id="a5231-746">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-747">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-747">- Content</span></span><br><span data-ttu-id="a5231-748">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-748">
         - TaskPane</span></span><br><span data-ttu-id="a5231-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-750">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="a5231-753">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-753">- ActiveView</span></span><br><span data-ttu-id="a5231-754">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-754">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-755">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-755">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-756">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-756">
         - File</span></span><br><span data-ttu-id="a5231-757">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-757">
         - PdfFile</span></span><br><span data-ttu-id="a5231-758">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-758">
         - Selection</span></span><br><span data-ttu-id="a5231-759">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-759">
         - Settings</span></span><br><span data-ttu-id="a5231-760">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-760">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-761">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-761">Office 2019 on Windows</span></span><br><span data-ttu-id="a5231-762">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-762">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-763">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-763">- Content</span></span><br><span data-ttu-id="a5231-764">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-764">
         - TaskPane</span></span><br><span data-ttu-id="a5231-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-766">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-767">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-768">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-768">- ActiveView</span></span><br><span data-ttu-id="a5231-769">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-769">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-770">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-770">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-771">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-771">
         - File</span></span><br><span data-ttu-id="a5231-772">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-772">
         - PdfFile</span></span><br><span data-ttu-id="a5231-773">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-773">
         - Selection</span></span><br><span data-ttu-id="a5231-774">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-774">
         - Settings</span></span><br><span data-ttu-id="a5231-775">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-775">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-776">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-776">Office 2016 on Windows</span></span><br><span data-ttu-id="a5231-777">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-777">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-778">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-778">- Content</span></span><br><span data-ttu-id="a5231-779">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-779">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a5231-780">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="a5231-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-782">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-782">- ActiveView</span></span><br><span data-ttu-id="a5231-783">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-783">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-784">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-784">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-785">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-785">
         - File</span></span><br><span data-ttu-id="a5231-786">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-786">
         - PdfFile</span></span><br><span data-ttu-id="a5231-787">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-787">
         - Selection</span></span><br><span data-ttu-id="a5231-788">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-788">
         - Settings</span></span><br><span data-ttu-id="a5231-789">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-789">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-790">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a5231-790">Office 2013 on Windows</span></span><br><span data-ttu-id="a5231-791">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-791">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-792">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-792">- Content</span></span><br><span data-ttu-id="a5231-793">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-793">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="a5231-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a5231-794">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="a5231-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-795">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-796">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-796">- ActiveView</span></span><br><span data-ttu-id="a5231-797">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-797">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-798">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-798">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-799">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-799">
         - File</span></span><br><span data-ttu-id="a5231-800">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-800">
         - PdfFile</span></span><br><span data-ttu-id="a5231-801">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-801">
         - Selection</span></span><br><span data-ttu-id="a5231-802">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-802">
         - Settings</span></span><br><span data-ttu-id="a5231-803">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-803">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-804">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-804">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="a5231-805">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-805">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-806">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-806">- Content</span></span><br><span data-ttu-id="a5231-807">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-807">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-808">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-809">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-810">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-810">- ActiveView</span></span><br><span data-ttu-id="a5231-811">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-811">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-812">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-812">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-813">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-813">
         - File</span></span><br><span data-ttu-id="a5231-814">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-814">
         - PdfFile</span></span><br><span data-ttu-id="a5231-815">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-815">
         - Selection</span></span><br><span data-ttu-id="a5231-816">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-816">
         - Settings</span></span><br><span data-ttu-id="a5231-817">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-817">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-818">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="a5231-818">Office apps on Mac</span></span><br><span data-ttu-id="a5231-819">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="a5231-819">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="a5231-820">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-820">- Content</span></span><br><span data-ttu-id="a5231-821">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-821">
         - TaskPane</span></span><br><span data-ttu-id="a5231-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-822">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-823">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="a5231-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="a5231-825">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="a5231-826">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-826">- ActiveView</span></span><br><span data-ttu-id="a5231-827">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-827">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-828">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-828">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-829">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-829">
         - File</span></span><br><span data-ttu-id="a5231-830">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-830">
         - PdfFile</span></span><br><span data-ttu-id="a5231-831">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-831">
         - Selection</span></span><br><span data-ttu-id="a5231-832">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-832">
         - Settings</span></span><br><span data-ttu-id="a5231-833">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-833">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-834">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-834">Office 2019 for Mac</span></span><br><span data-ttu-id="a5231-835">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-835">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-836">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-836">- Content</span></span><br><span data-ttu-id="a5231-837">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-837">
         - TaskPane</span></span><br><span data-ttu-id="a5231-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-838">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-839">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-841">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-841">- ActiveView</span></span><br><span data-ttu-id="a5231-842">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-842">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-843">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-843">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-844">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-844">
         - File</span></span><br><span data-ttu-id="a5231-845">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-845">
         - PdfFile</span></span><br><span data-ttu-id="a5231-846">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-846">
         - Selection</span></span><br><span data-ttu-id="a5231-847">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-847">
         - Settings</span></span><br><span data-ttu-id="a5231-848">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-848">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-849">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-849">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="a5231-850">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-850">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-851">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-851">- Content</span></span><br><span data-ttu-id="a5231-852">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-852">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="a5231-853">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="a5231-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-855">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="a5231-855">- ActiveView</span></span><br><span data-ttu-id="a5231-856">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="a5231-856">
         - CompressedFile</span></span><br><span data-ttu-id="a5231-857">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-857">
         - DocumentEvents</span></span><br><span data-ttu-id="a5231-858">
         - File</span><span class="sxs-lookup"><span data-stu-id="a5231-858">
         - File</span></span><br><span data-ttu-id="a5231-859">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="a5231-859">
         - PdfFile</span></span><br><span data-ttu-id="a5231-860">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-860">
         - Selection</span></span><br><span data-ttu-id="a5231-861">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-861">
         - Settings</span></span><br><span data-ttu-id="a5231-862">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-862">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="a5231-863">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="a5231-863">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="a5231-864">OneNote</span><span class="sxs-lookup"><span data-stu-id="a5231-864">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a5231-865">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-865">Platform</span></span></th>
    <th><span data-ttu-id="a5231-866">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-866">Extension points</span></span></th>
    <th><span data-ttu-id="a5231-867">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-867">API requirement sets</span></span></th>
    <th><span data-ttu-id="a5231-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-868"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-869">Office on the web</span><span class="sxs-lookup"><span data-stu-id="a5231-869">Office on the web</span></span></td>
    <td> <span data-ttu-id="a5231-870">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="a5231-870">- Content</span></span><br><span data-ttu-id="a5231-871">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-871">
         - TaskPane</span></span><br><span data-ttu-id="a5231-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="a5231-872">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="a5231-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-873">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="a5231-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-874">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="a5231-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-876">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="a5231-876">- DocumentEvents</span></span><br><span data-ttu-id="a5231-877">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-877">
         - HtmlCoercion</span></span><br><span data-ttu-id="a5231-878">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="a5231-878">
         - Settings</span></span><br><span data-ttu-id="a5231-879">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-879">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="a5231-880">Project</span><span class="sxs-lookup"><span data-stu-id="a5231-880">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="a5231-881">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="a5231-881">Platform</span></span></th>
    <th><span data-ttu-id="a5231-882">拡張点</span><span class="sxs-lookup"><span data-stu-id="a5231-882">Extension points</span></span></th>
    <th><span data-ttu-id="a5231-883">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-883">API requirement sets</span></span></th>
    <th><span data-ttu-id="a5231-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="a5231-884"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-885">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="a5231-885">Office 2019 on Windows</span></span><br><span data-ttu-id="a5231-886">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-886">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-887">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-887">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-888">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-889">- Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-889">- Selection</span></span><br><span data-ttu-id="a5231-890">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-890">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-891">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="a5231-891">Office 2016 on Windows</span></span><br><span data-ttu-id="a5231-892">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-892">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-893">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-893">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-894">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-895">- Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-895">- Selection</span></span><br><span data-ttu-id="a5231-896">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-896">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="a5231-897">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a5231-897">Office 2013 on Windows</span></span><br><span data-ttu-id="a5231-898">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="a5231-898">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="a5231-899">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="a5231-899">- TaskPane</span></span></td>
    <td> <span data-ttu-id="a5231-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="a5231-900">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="a5231-901">- Selection</span><span class="sxs-lookup"><span data-stu-id="a5231-901">- Selection</span></span><br><span data-ttu-id="a5231-902">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="a5231-902">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="a5231-903">関連項目</span><span class="sxs-lookup"><span data-stu-id="a5231-903">See also</span></span>

- [<span data-ttu-id="a5231-904">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="a5231-904">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="a5231-905">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-905">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a5231-906">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-906">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="a5231-907">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="a5231-907">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="a5231-908">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="a5231-908">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="a5231-909">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="a5231-909">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="a5231-910">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="a5231-910">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="a5231-911">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="a5231-911">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="a5231-912">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="a5231-912">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="a5231-913">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="a5231-913">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="a5231-914">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="a5231-914">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
