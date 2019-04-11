---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、Word、Outlook、PowerPoint、OneNote、および Project のサポートされる要件セット。
ms.date: 04/03/2019
localization_priority: Priority
ms.openlocfilehash: a9ecd44edf9221a403eb42756cd1e9f5e676ad01
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477594"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="f2cb9-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="f2cb9-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="f2cb9-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-106">The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="f2cb9-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="f2cb9-108">Excel</span><span class="sxs-lookup"><span data-stu-id="f2cb9-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f2cb9-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f2cb9-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="f2cb9-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f2cb9-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-111">API requirement sets</span></span></th>
    <th style="width:40%"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="f2cb9-112">共通 API</span><span class="sxs-lookup"><span data-stu-id="f2cb9-112">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-113">Office Online</span><span class="sxs-lookup"><span data-stu-id="f2cb9-113">Office Online</span></span></td>
    <td> - <span data-ttu-id="f2cb9-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-114">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-115">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-115">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-116">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-116">add-in commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-117">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-117">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-118">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-118">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-119">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-119">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-120">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-120">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-121">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-121">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-122">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-122">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-123">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-123">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-124">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="f2cb9-124">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-125">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-125">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="f2cb9-126">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-126">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-127">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-127">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-128">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-128">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-129">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-129">File</span></span><br>
        - <span data-ttu-id="f2cb9-130">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-130">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-131">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-131">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-132">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-132">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-133">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-133">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-134">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-134">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-135">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-135">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-136">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-136">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-137">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-137">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-138">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-138">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-139">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-139">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-140">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-140">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-141">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-141">add-in commands</span></span></a>
    </td>
    <td>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-142">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-142">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-143">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-143">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-144">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-144">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-145">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-145">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-146">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-146">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-147">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-147">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-148">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-148">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-149">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="f2cb9-149">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-150">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-150">DialogApi 1.1</span></span></a></td>
    <td>
        - <span data-ttu-id="f2cb9-151">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-151">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-152">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-152">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-153">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-153">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-154">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-154">File</span></span><br>
        - <span data-ttu-id="f2cb9-155">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-155">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-156">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-156">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-157">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-157">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-158">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-158">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-159">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-159">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-160">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-160">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-161">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-161">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-162">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-162">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-163">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-163">Office 2019 for Windows</span></span></td>
    <td>- <span data-ttu-id="f2cb9-164">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-164"> Taskpane</span></span><br>
        - <span data-ttu-id="f2cb9-165">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-165">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-166">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-166">add-in commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-167">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-167">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-168">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-168">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-169">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-169">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-170">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-170">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-171">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-171">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-172">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-172">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-173">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-173">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-174">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="f2cb9-174">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-175">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-175">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="f2cb9-176">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-176">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-177">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-177">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-178">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-178">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-179">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-179">File</span></span><br>
        - <span data-ttu-id="f2cb9-180">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-180">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-181">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-181">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-182">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-182">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-183">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-183">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-184">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-184">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-185">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-185">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-186">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-186">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-187">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-187">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-188">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-188">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-189">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-189">Office 2016 for Windows</span></span></td>
    <td>- <span data-ttu-id="f2cb9-190">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-190">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-191">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-191">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-192">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-192">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-193">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-193">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="f2cb9-194">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-194">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-195">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-195">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-196">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-196">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-197">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-197">File</span></span><br>
        - <span data-ttu-id="f2cb9-198">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-198">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-199">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-199">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-200">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-200">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-201">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-201">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-202">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-202">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-203">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-203">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-204">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-204">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-205">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-205">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-206">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-206">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-207">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-207">Office 2013 for Windows</span></span></td>
    <td>
        - <span data-ttu-id="f2cb9-208">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-208">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-209">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-209">Content</span></span></td>
    <td>  - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-210">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-210">DialogApi 1.1</span></span></a>*</td>
    <td>
        - <span data-ttu-id="f2cb9-211">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-211">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-212">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-212">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-213">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-213">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-214">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-214">File</span></span><br>
        - <span data-ttu-id="f2cb9-215">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-215">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-216">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-216">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-217">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-217">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-218">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-218">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-219">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-219">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-220">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-220">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-221">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-221">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-222">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-222">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-223">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-223">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-224">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="f2cb9-224">Office 365 for iPad</span></span></td>
    <td>- <span data-ttu-id="f2cb9-225">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-225">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-226">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-226">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-227">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-227">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-228">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-228">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-229">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-229">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-230">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-230">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-231">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-231">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-232">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-232">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-233">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-233">ExcelApi 1.7</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-234">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="f2cb9-234">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-235">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-235">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="f2cb9-236">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-236">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-237">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-237">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-238">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-238">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-239">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-239">File</span></span><br>
        - <span data-ttu-id="f2cb9-240">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-240">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-241">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-241">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-242">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-242">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-243">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-243">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-244">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-244">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-245">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-245">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-246">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-246">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-247">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-247">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-248">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-248">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-249">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-249">Office 365 for Mac</span></span></td>
    <td>- <span data-ttu-id="f2cb9-250">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-250">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-251">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-251">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-252">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-252">add-in commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-253">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-253">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-254">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-254">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-255">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-255">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-256">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-256">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-257">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-257">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-258">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-258">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-259">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-259">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-260">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="f2cb9-260">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-261">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-261">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="f2cb9-262">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-262">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-263">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-263">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-264">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-264">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-265">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-265">File</span></span><br>
        - <span data-ttu-id="f2cb9-266">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-266">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-267">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-267">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-268">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-268">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-269">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-269">PdfFile</span></span><br>
        - <span data-ttu-id="f2cb9-270">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-270">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-271">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-271">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-272">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-272">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-273">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-273">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-274">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-274">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-275">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-275">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-276">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-276">Office 2019 for Mac</span></span></td>
    <td>- <span data-ttu-id="f2cb9-277">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-277">TaskPane</span></span><br>
        - <span data-ttu-id="f2cb9-278">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-278">Content</span></span><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-279">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-279">add-in commands</span></span></a></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-280">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-280">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-281">ExcelApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-281">ExcelApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-282">ExcelApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-282">ExcelApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-283">ExcelApi 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-283">ExcelApi 1.4</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-284">ExcelApi 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-284">ExcelApi 1.5</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-285">ExcelApi 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-285">ExcelApi 1.6</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-286">ExcelApi 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-286">ExcelApi 1.7</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-287">ExcelApi 1.8</span><span class="sxs-lookup"><span data-stu-id="f2cb9-287">ExcelApi 1.8</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-288">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-288">DialogApi 1.1</span></span></a></td>
    <td>- <span data-ttu-id="f2cb9-289">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-289">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-290">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-290">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-291">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-291">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-292">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-292">File</span></span><br>
        - <span data-ttu-id="f2cb9-293">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-293">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-294">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-294">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-295">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-295">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-296">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-296">PdfFile</span></span><br>
        - <span data-ttu-id="f2cb9-297">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-297">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-298">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-298">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-299">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-299">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-300">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-300">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-301">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-301">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-302">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-302">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-303">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-303">Office 2016 for Mac</span></span></td>
    <td>- <span data-ttu-id="f2cb9-304">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-304"> Taskpane</span></span><br>
        - <span data-ttu-id="f2cb9-305">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-305">Content</span></span></td>
    <td>- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets"><span data-ttu-id="f2cb9-306">ExcelApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-306">ExcelApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-307">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-307">DialogApi 1.1</span></span></a>*</td>
    <td>- <span data-ttu-id="f2cb9-308">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-308">BindingEvents</span></span><br>
        - <span data-ttu-id="f2cb9-309">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-309">CompressedFile</span></span><br>
        - <span data-ttu-id="f2cb9-310">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-310">DocumentEvents</span></span><br>
        - <span data-ttu-id="f2cb9-311">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-311">File</span></span><br>
        - <span data-ttu-id="f2cb9-312">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-312">ImageCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-313">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-313">MatrixBindings</span></span><br>
        - <span data-ttu-id="f2cb9-314">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-314">MatrixCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-315">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-315">PdfFile</span></span><br>
        - <span data-ttu-id="f2cb9-316">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-316">Selection</span></span><br>
        - <span data-ttu-id="f2cb9-317">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-317">Settings</span></span><br>
        - <span data-ttu-id="f2cb9-318">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-318">TableBindings</span></span><br>
        - <span data-ttu-id="f2cb9-319">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-319">TableCoercion</span></span><br>
        - <span data-ttu-id="f2cb9-320">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-320">TextBindings</span></span><br>
        - <span data-ttu-id="f2cb9-321">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-321">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="f2cb9-322">&ast; - リリース後の更新プログラムで追加されました。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-322">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="outlook"></a><span data-ttu-id="f2cb9-323">Outlook</span><span class="sxs-lookup"><span data-stu-id="f2cb9-323">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cb9-324">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-324">Platform</span></span></th>
    <th><span data-ttu-id="f2cb9-325">拡張点</span><span class="sxs-lookup"><span data-stu-id="f2cb9-325">Extension points</span></span></th>
    <th><span data-ttu-id="f2cb9-326">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-326">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="f2cb9-327">共通 API</span><span class="sxs-lookup"><span data-stu-id="f2cb9-327">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-328">Office Online</span><span class="sxs-lookup"><span data-stu-id="f2cb9-328">Office Online</span></span></td>
    <td> - <span data-ttu-id="f2cb9-329">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-329">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-330">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-330">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-331">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-331">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-332">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-332">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-333">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-333">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-334">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-334">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-335">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-335">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-336">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-336">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="f2cb9-337">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-337">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="f2cb9-338">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-338">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="f2cb9-339">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-339">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-340">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-340">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-341">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-341">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-342">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-342">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-343">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-343">add-in commands</span></span></a><br>
      - <span data-ttu-id="f2cb9-344">モジュール</span><span class="sxs-lookup"><span data-stu-id="f2cb9-344">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-345">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-345">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-346">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-346">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-347">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-347">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-348">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-348">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-349">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-349">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="f2cb9-350">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-350">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="f2cb9-351">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-351">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="f2cb9-352">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-352">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-353">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-353">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-354">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-354">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-355">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-355">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-356">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-356">add-in commands</span></span></a><br>
      - <span data-ttu-id="f2cb9-357">モジュール</span><span class="sxs-lookup"><span data-stu-id="f2cb9-357">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-358">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-358">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-359">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-359">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-360">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-360">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-361">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-361">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-362">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-362">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="f2cb9-363">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-363">Mailbox 1.6</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7"><span data-ttu-id="f2cb9-364">Mailbox 1.7</span><span class="sxs-lookup"><span data-stu-id="f2cb9-364">Mailbox 1.7</span></span></a></td>
    <td><span data-ttu-id="f2cb9-365">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-365">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-366">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-366">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-367">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-367">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-368">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-368">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-369">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-369">add-in commands</span></span></a><br>
      - <span data-ttu-id="f2cb9-370">モジュール</span><span class="sxs-lookup"><span data-stu-id="f2cb9-370">Modules</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-371">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-371">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-372">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-372">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-373">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-373">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-374">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-374">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="f2cb9-375">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-375">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-376">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-376">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-377">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-377">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-378">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-378">Mail Compose</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-379">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-379">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-380">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-380">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-381">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-381">Mailbox 1.3</span></span></a>*<br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-382">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-382">Mailbox 1.4</span></span></a>*</td>
    <td><span data-ttu-id="f2cb9-383">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-383">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-384">Office 365 for iOS</span><span class="sxs-lookup"><span data-stu-id="f2cb9-384">Office 365 for iOS</span></span></td>
    <td> - <span data-ttu-id="f2cb9-385">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-385">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-386">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-386">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-387">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-387">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-388">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-388">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-389">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-389">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-390">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-390">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-391">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-391">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="f2cb9-392">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-392">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-393">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-393">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-394">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-394">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-395">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-395">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-396">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-396">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-397">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-397">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-398">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-398">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-399">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-399">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-400">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-400">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-401">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-401">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="f2cb9-402">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-402">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="f2cb9-403">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-403">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-404">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-404">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-405">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-405">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-406">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-406">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-407">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-407">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-408">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-408">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-409">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-409">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-410">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-410">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-411">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-411">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-412">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-412">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="f2cb9-413">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-413">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="f2cb9-414">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-414">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-415">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-415">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-416">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-416">Mail Read</span></span><br>
      - <span data-ttu-id="f2cb9-417">メールの作成</span><span class="sxs-lookup"><span data-stu-id="f2cb9-417">Mail Compose</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-418">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-418">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-419">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-419">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-420">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-420">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-421">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-421">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-422">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-422">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-423">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-423">Mailbox 1.5</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6"><span data-ttu-id="f2cb9-424">Mailbox 1.6</span><span class="sxs-lookup"><span data-stu-id="f2cb9-424">Mailbox 1.6</span></span></a></td>
    <td><span data-ttu-id="f2cb9-425">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-425">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-426">Office 365 for Android</span><span class="sxs-lookup"><span data-stu-id="f2cb9-426">Office 365 for Android</span></span></td>
    <td> - <span data-ttu-id="f2cb9-427">メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="f2cb9-427">Mail Read</span></span><br>
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-428">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-428">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1"><span data-ttu-id="f2cb9-429">Mailbox 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-429">Mailbox 1.1</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2"><span data-ttu-id="f2cb9-430">Mailbox 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-430">Mailbox 1.2</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3"><span data-ttu-id="f2cb9-431">Mailbox 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-431">Mailbox 1.3</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4"><span data-ttu-id="f2cb9-432">Mailbox 1.4</span><span class="sxs-lookup"><span data-stu-id="f2cb9-432">Mailbox 1.4</span></span></a><br>
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5"><span data-ttu-id="f2cb9-433">Mailbox 1.5</span><span class="sxs-lookup"><span data-stu-id="f2cb9-433">Mailbox 1.5</span></span></a></td>
    <td><span data-ttu-id="f2cb9-434">使用不可</span><span class="sxs-lookup"><span data-stu-id="f2cb9-434">Not available</span></span></td>
  </tr>
</table>

*<span data-ttu-id="f2cb9-435">&ast; - リリース後の更新プログラムで追加されました。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-435">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="word"></a><span data-ttu-id="f2cb9-436">Word</span><span class="sxs-lookup"><span data-stu-id="f2cb9-436">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cb9-437">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-437">Platform</span></span></th>
    <th><span data-ttu-id="f2cb9-438">拡張点</span><span class="sxs-lookup"><span data-stu-id="f2cb9-438">Extension points</span></span></th>
    <th><span data-ttu-id="f2cb9-439">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-439">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="f2cb9-440">共通 API</span><span class="sxs-lookup"><span data-stu-id="f2cb9-440">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-441">Office Online</span><span class="sxs-lookup"><span data-stu-id="f2cb9-441">Office Online</span></span></td>
    <td> - <span data-ttu-id="f2cb9-442">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-442">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-443">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-443">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-444">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-444">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-445">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-445">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-446">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-446">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-447">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-447">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-448">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-448">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-449">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-449">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-450">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-450">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-451">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-451">File</span></span><br>
         - <span data-ttu-id="f2cb9-452">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-452">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-453">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-453">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-454">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-454">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-455">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-455">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-456">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-456">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-457">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-457">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-458">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-458">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-459">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-459">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-460">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-460">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-461">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-461">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-462">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-462">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-463">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-463">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-464">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-464">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-465">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-465">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-466">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-466">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-467">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-467">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-468">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-468">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-469">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-469">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-470">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-470">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-471">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-471">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-472">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-472">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-473">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-473">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-474">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-474">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-475">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-475">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-476">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-476">File</span></span><br>
         - <span data-ttu-id="f2cb9-477">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-477">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-478">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-478">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-479">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-479">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-480">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-480">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-481">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-481">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-482">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-482">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-483">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-483">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-484">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-484">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-485">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-485">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-486">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-486">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-487">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-487">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-488">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-488">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-489">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-489">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-490">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-490">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-491">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-491"> Taskpane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-492">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-492">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-493">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-493">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-494">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-494">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-495">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-495">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-496">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-496">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-497">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-497">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-498">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-498">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-499">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-499">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-500">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-500">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-501">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-501">File</span></span><br>
         - <span data-ttu-id="f2cb9-502">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-502">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-503">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-503">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-504">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-504">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-505">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-505">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-506">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-506">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-507">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-507">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-508">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-508">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-509">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-509">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-510">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-510">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-511">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-511">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-512">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-512">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-513">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-513">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-514">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-514">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-515">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-515">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-516">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-516"> Taskpane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-517">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-517">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-518">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-518">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="f2cb9-519">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-519">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-520">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-520">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-521">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-521">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-522">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-522">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-523">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-523">File</span></span><br>
         - <span data-ttu-id="f2cb9-524">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-524">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-525">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-525">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-526">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-526">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-527">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-527">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-528">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-528">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-529">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-529">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-530">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-530">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-531">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-531">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-532">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-532">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-533">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-533">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-534">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-534">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-535">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-535">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-536">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-536">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-537">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-537">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-538">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-538">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-539">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-539">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="f2cb9-540">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-540">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-541">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-541">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-542">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-542">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-543">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-543">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-544">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-544">File</span></span><br>
         - <span data-ttu-id="f2cb9-545">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-545">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-546">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-546">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-547">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-547">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-548">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-548">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-549">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-549">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-550">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-550">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-551">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-551">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-552">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-552">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-553">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-553">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-554">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-554">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-555">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-555">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-556">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-556">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-557">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-557">TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-558">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="f2cb9-558">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="f2cb9-559">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-559">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-560">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-560">WordApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-561">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-561">WordApi 1.2</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-562">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-562">WordApi 1.3</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-563">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-563">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="f2cb9-564">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-564">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-565">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-565">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-566">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-566">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-567">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-567">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-568">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-568">File</span></span><br>
         - <span data-ttu-id="f2cb9-569">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-569">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-570">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-570">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-571">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-571">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-572">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-572">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-573">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-573">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-574">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-574">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-575">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-575">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-576">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-576">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-577">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-577">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-578">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-578">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-579">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-579">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-580">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-580">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-581">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-581">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-582">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-582">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-583">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-583">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-584">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-584">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-585">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-585">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-586">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-586">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-587">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-587">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-588">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-588">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="f2cb9-589">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-589">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-590">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-590">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-591">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-591">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-592">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-592">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-593">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-593">File</span></span><br>
         - <span data-ttu-id="f2cb9-594">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-594">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-595">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-595">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-596">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-596">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-597">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-597">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-598">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-598">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-599">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-599">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-600">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-600">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-601">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-601">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-602">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-602">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-603">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-603">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-604">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-604">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-605">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-605">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-606">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-606">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-607">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-607">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-608">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-608">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-609">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-609">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-610">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-610">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-611">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="f2cb9-611">WordApi 1.2</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-612">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="f2cb9-612">WordApi 1.3</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-613">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-613">DialogApi 1.1</span></span></a>
</td>
    <td> - <span data-ttu-id="f2cb9-614">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-614">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-615">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-615">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-616">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-616">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-617">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-617">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-618">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-618">File</span></span><br>
         - <span data-ttu-id="f2cb9-619">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-619">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-620">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-620">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-621">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-621">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-622">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-622">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-623">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-623">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-624">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-624">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-625">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-625">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-626">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-626">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-627">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-627">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-628">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-628">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-629">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-629">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-630">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-630">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-631">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-631">TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-632">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-632">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-633">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-633"> Taskpane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets"><span data-ttu-id="f2cb9-634">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-634">WordApi 1.1</span></span></a><br>
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-635">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-635">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="f2cb9-636">BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-636">BindingEvents</span></span><br>
         - <span data-ttu-id="f2cb9-637">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-637">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-638">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f2cb9-638">CustomXmlParts</span></span><br>
         - <span data-ttu-id="f2cb9-639">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-639">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-640">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-640">File</span></span><br>
         - <span data-ttu-id="f2cb9-641">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-641">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-642">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-642">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-643">MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-643">MatrixBindings</span></span><br>
         - <span data-ttu-id="f2cb9-644">MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-644">MatrixCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-645">OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-645">OoxmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-646">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-646">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-647">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-647">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-648">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-648">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-649">TableBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-649">TableBindings</span></span><br>
         - <span data-ttu-id="f2cb9-650">TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-650">TableCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-651">TextBindings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-651">TextBindings</span></span><br>
         - <span data-ttu-id="f2cb9-652">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-652">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-653">TextFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-653">TextFile</span></span> </td>
  </tr>
</table>

*<span data-ttu-id="f2cb9-654">&ast; - リリース後の更新プログラムで追加されました。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-654">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="powerpoint"></a><span data-ttu-id="f2cb9-655">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f2cb9-655">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cb9-656">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-656">Platform</span></span></th>
    <th><span data-ttu-id="f2cb9-657">拡張点</span><span class="sxs-lookup"><span data-stu-id="f2cb9-657">Extension points</span></span></th>
    <th><span data-ttu-id="f2cb9-658">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-658">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="f2cb9-659">共通 API</span><span class="sxs-lookup"><span data-stu-id="f2cb9-659">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-660">Office Online</span><span class="sxs-lookup"><span data-stu-id="f2cb9-660">Office Online</span></span></td>
    <td> - <span data-ttu-id="f2cb9-661">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-661">Content</span></span><br>
         - <span data-ttu-id="f2cb9-662">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-662">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-663">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-663">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-664">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-664">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-665">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-665">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-666">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-666">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-667">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-667">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-668">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-668">File</span></span><br>
         - <span data-ttu-id="f2cb9-669">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-669">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-670">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-670">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-671">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-671">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-672">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-672">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-673">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-673">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-674">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-674">Office 365 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-675">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-675">Content</span></span><br>
         - <span data-ttu-id="f2cb9-676">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-676">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-677">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-677">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-678">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-678">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-679">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-679">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-680">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-680">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-681">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-681">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-682">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-682">File</span></span><br>
         - <span data-ttu-id="f2cb9-683">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-683">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-684">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-684">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-685">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-685">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-686">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-686">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-687">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-687">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-688">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-688">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-689">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-689">Content</span></span><br>
         - <span data-ttu-id="f2cb9-690">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-690">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-691">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-691">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-692">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-692">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-693">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-693">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-694">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-694">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-695">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-695">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-696">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-696">File</span></span><br>
         - <span data-ttu-id="f2cb9-697">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-697">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-698">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-698">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-699">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-699">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-700">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-700">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-701">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-701">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-702">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-702">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-703">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-703">Content</span></span><br>
         - <span data-ttu-id="f2cb9-704">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-704">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-705">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-705">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="f2cb9-706">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-706">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-707">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-707">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-708">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-708">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-709">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-709">File</span></span><br>
         - <span data-ttu-id="f2cb9-710">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-710">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-711">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-711">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-712">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-712">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-713">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-713">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-714">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-714">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-715">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-715">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-716">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-716">Content</span></span><br>
         - <span data-ttu-id="f2cb9-717">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-717">TaskPane</span></span><br>
    </td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-718">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-718">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="f2cb9-719">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-719">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-720">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-720">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-721">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-721">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-722">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-722">File</span></span><br>
         - <span data-ttu-id="f2cb9-723">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-723">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-724">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-724">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-725">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-725">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-726">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-726">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-727">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-727">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-728">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="f2cb9-728">Office 365 for iPad</span></span></td>
    <td> - <span data-ttu-id="f2cb9-729">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-729">Content</span></span><br>
         - <span data-ttu-id="f2cb9-730">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-730">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-731">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-731">DialogApi 1.1</span></span></a></td>
     <td> - <span data-ttu-id="f2cb9-732">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-732">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-733">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-733">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-734">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-734">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-735">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-735">File</span></span><br>
         - <span data-ttu-id="f2cb9-736">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-736">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-737">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-737">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-738">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-738">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-739">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-739">TextCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-740">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-740">ImageCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-741">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-741">Office 365 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-742">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-742">Content</span></span><br>
         - <span data-ttu-id="f2cb9-743">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-743">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-744">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-744">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-745">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-745">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-746">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-746">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-747">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-747">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-748">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-748">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-749">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-749">File</span></span><br>
         - <span data-ttu-id="f2cb9-750">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-750">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-751">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-751">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-752">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-752">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-753">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-753">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-754">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-754">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-755">Office 2019 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-755">Office 2019 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-756">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-756">Content</span></span><br>
         - <span data-ttu-id="f2cb9-757">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-757">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-758">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-758">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-759">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-759">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-760">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-760">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-761">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-761">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-762">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-762">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-763">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-763">File</span></span><br>
         - <span data-ttu-id="f2cb9-764">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-764">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-765">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-765">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-766">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-766">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-767">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-767">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-768">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-768">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-769">Office 2016 for Mac</span><span class="sxs-lookup"><span data-stu-id="f2cb9-769">Office 2016 for Mac</span></span></td>
    <td> - <span data-ttu-id="f2cb9-770">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-770">Content</span></span><br>
         - <span data-ttu-id="f2cb9-771">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-771">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-772">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-772">DialogApi 1.1</span></span></a>*</td>
    <td> - <span data-ttu-id="f2cb9-773">ActiveView</span><span class="sxs-lookup"><span data-stu-id="f2cb9-773">ActiveView</span></span><br>
         - <span data-ttu-id="f2cb9-774">CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-774">CompressedFile</span></span><br>
         - <span data-ttu-id="f2cb9-775">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-775">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-776">File</span><span class="sxs-lookup"><span data-stu-id="f2cb9-776">File</span></span><br>
         - <span data-ttu-id="f2cb9-777">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-777">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-778">PdfFile</span><span class="sxs-lookup"><span data-stu-id="f2cb9-778">PdfFile</span></span><br>
         - <span data-ttu-id="f2cb9-779">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-779">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-780">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-780">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-781">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-781">TextCoercion</span></span></td>
  </tr>
</table>

*<span data-ttu-id="f2cb9-782">&ast; - リリース後の更新プログラムで追加されました。</span><span class="sxs-lookup"><span data-stu-id="f2cb9-782">&ast; - Added with post-release updates.</span></span>*

<br/>

## <a name="onenote"></a><span data-ttu-id="f2cb9-783">OneNote</span><span class="sxs-lookup"><span data-stu-id="f2cb9-783">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cb9-784">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-784">Platform</span></span></th>
    <th><span data-ttu-id="f2cb9-785">拡張点</span><span class="sxs-lookup"><span data-stu-id="f2cb9-785">Extension points</span></span></th>
    <th><span data-ttu-id="f2cb9-786">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-786">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="f2cb9-787">共通 API</span><span class="sxs-lookup"><span data-stu-id="f2cb9-787">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-788">Office Online</span><span class="sxs-lookup"><span data-stu-id="f2cb9-788">Office Online</span></span></td>
    <td> - <span data-ttu-id="f2cb9-789">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2cb9-789">Content</span></span><br>
         - <span data-ttu-id="f2cb9-790">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-790">TaskPane</span></span><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets"><span data-ttu-id="f2cb9-791">アドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="f2cb9-791">add-in commands</span></span></a></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets"><span data-ttu-id="f2cb9-792">OneNoteApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-792">OneNoteApi 1.1</span></span></a><br>
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-793">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-793">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-794">DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f2cb9-794">DocumentEvents</span></span><br>
         - <span data-ttu-id="f2cb9-795">HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-795">HtmlCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-796">ImageCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-796">ImageCoercion</span></span><br>
         - <span data-ttu-id="f2cb9-797">Settings</span><span class="sxs-lookup"><span data-stu-id="f2cb9-797">Settings</span></span><br>
         - <span data-ttu-id="f2cb9-798">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-798">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="f2cb9-799">Project</span><span class="sxs-lookup"><span data-stu-id="f2cb9-799">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f2cb9-800">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-800">Platform</span></span></th>
    <th><span data-ttu-id="f2cb9-801">拡張点</span><span class="sxs-lookup"><span data-stu-id="f2cb9-801">Extension points</span></span></th>
    <th><span data-ttu-id="f2cb9-802">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-802">API requirement sets</span></span></th>
    <th><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b><span data-ttu-id="f2cb9-803">共通 API</span><span class="sxs-lookup"><span data-stu-id="f2cb9-803">Common APIs</span></span></b></a></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-804">Office 2019 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-804">Office 2019 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-805">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-805">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-806">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-806">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-807">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-807">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-808">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-808">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-809">Office 2016 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-809">Office 2016 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-810">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-810">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-811">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-811">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-812">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-812">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-813">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-813">TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f2cb9-814">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="f2cb9-814">Office 2013 for Windows</span></span></td>
    <td> - <span data-ttu-id="f2cb9-815">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2cb9-815">TaskPane</span></span></td>
    <td> - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets"><span data-ttu-id="f2cb9-816">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="f2cb9-816">DialogApi 1.1</span></span></a></td>
    <td> - <span data-ttu-id="f2cb9-817">Selection</span><span class="sxs-lookup"><span data-stu-id="f2cb9-817">Selection</span></span><br>
         - <span data-ttu-id="f2cb9-818">TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f2cb9-818">TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="f2cb9-819">関連項目</span><span class="sxs-lookup"><span data-stu-id="f2cb9-819">See also</span></span>

- [<span data-ttu-id="f2cb9-820">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="f2cb9-820">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="f2cb9-821">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-821">Common API requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
- [<span data-ttu-id="f2cb9-822">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="f2cb9-822">Add-in Commands requirement sets</span></span>](/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets)
- [<span data-ttu-id="f2cb9-823">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="f2cb9-823">JavaScript API for Office reference</span></span>](/office/dev/add-ins/reference/javascript-api-for-office)
- [<span data-ttu-id="f2cb9-824">Office 2016 C2R および Office 2019 の更新履歴</span><span class="sxs-lookup"><span data-stu-id="f2cb9-824">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="f2cb9-825">Office 2013 の更新履歴</span><span class="sxs-lookup"><span data-stu-id="f2cb9-825">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="f2cb9-826">Windows インストーラー (MSI) を使用しているバージョンの Office の最新の更新プログラム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-826">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="f2cb9-827">Windows インストーラー (MSI) を使用しているバージョンの Outlook の最新の更新プログラム</span><span class="sxs-lookup"><span data-stu-id="f2cb9-827">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)