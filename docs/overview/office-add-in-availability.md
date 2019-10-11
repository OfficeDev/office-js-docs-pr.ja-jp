---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 28d63866a03bcae99829d3a6b6c6198059a92bdc
ms.sourcegitcommit: 4d9f3e177b0bcd62804d5045f52b03e441af244f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/10/2019
ms.locfileid: "37440151"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="26bb8-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="26bb8-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="26bb8-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="26bb8-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="26bb8-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="26bb8-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="26bb8-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="26bb8-108">Excel</span><span class="sxs-lookup"><span data-stu-id="26bb8-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="26bb8-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="26bb8-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="26bb8-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="26bb8-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="26bb8-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-114">- TaskPane</span></span><br><span data-ttu-id="26bb8-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-115">
        - Content</span></span><br><span data-ttu-id="26bb8-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-116">
        - Custom Functions</span></span><br><span data-ttu-id="26bb8-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="26bb8-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="26bb8-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="26bb8-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="26bb8-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="26bb8-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="26bb8-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="26bb8-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="26bb8-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="26bb8-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="26bb8-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-128">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-128">
        - BindingEvents</span></span><br><span data-ttu-id="26bb8-129">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-129">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-130">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-130">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-131">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-131">
        - File</span></span><br><span data-ttu-id="26bb8-132">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-132">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-133">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-133">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-134">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-134">
        - Selection</span></span><br><span data-ttu-id="26bb8-135">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-135">
        - Settings</span></span><br><span data-ttu-id="26bb8-136">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-136">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-137">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-137">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-138">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-138">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-139">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-139">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-140">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-140">Office on Windows</span></span><br><span data-ttu-id="26bb8-141">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-141">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-142">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-142">- TaskPane</span></span><br><span data-ttu-id="26bb8-143">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-143">
        - Content</span></span><br><span data-ttu-id="26bb8-144">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-144">
        - Custom Functions</span></span><br><span data-ttu-id="26bb8-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="26bb8-145">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="26bb8-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-146">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-147">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="26bb8-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="26bb8-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="26bb8-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="26bb8-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="26bb8-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="26bb8-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="26bb8-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="26bb8-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="26bb8-158">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-158">
        - BindingEvents</span></span><br><span data-ttu-id="26bb8-159">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-159">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-160">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-160">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-161">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-161">
        - File</span></span><br><span data-ttu-id="26bb8-162">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-162">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-163">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-163">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-164">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-164">
        - Selection</span></span><br><span data-ttu-id="26bb8-165">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-165">
        - Settings</span></span><br><span data-ttu-id="26bb8-166">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-166">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-167">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-167">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-168">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-168">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-169">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-169">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-170">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-170">Office 2019 on Windows</span></span><br><span data-ttu-id="26bb8-171">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-171">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="26bb8-172">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-172">- TaskPane</span></span><br><span data-ttu-id="26bb8-173">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-173">
        - Content</span></span><br><span data-ttu-id="26bb8-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-174">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="26bb8-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-175">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-176">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="26bb8-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-177">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="26bb8-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-178">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="26bb8-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="26bb8-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-180">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="26bb8-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="26bb8-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="26bb8-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-185">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-185">- BindingEvents</span></span><br><span data-ttu-id="26bb8-186">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-186">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-187">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-187">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-188">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-188">
        - File</span></span><br><span data-ttu-id="26bb8-189">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-189">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-190">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-190">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-191">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-191">
        - Selection</span></span><br><span data-ttu-id="26bb8-192">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-192">
        - Settings</span></span><br><span data-ttu-id="26bb8-193">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-193">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-194">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-194">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-195">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-195">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-196">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-196">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-197">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-197">Office 2016 on Windows</span></span><br><span data-ttu-id="26bb8-198">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-198">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="26bb8-199">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-199">- TaskPane</span></span><br><span data-ttu-id="26bb8-200">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-200">
        - Content</span></span></td>
    <td><span data-ttu-id="26bb8-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-201">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-202">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="26bb8-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-203">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-204">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-204">- BindingEvents</span></span><br><span data-ttu-id="26bb8-205">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-205">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-206">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-206">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-207">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-207">
        - File</span></span><br><span data-ttu-id="26bb8-208">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-208">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-209">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-209">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-210">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-210">
        - Selection</span></span><br><span data-ttu-id="26bb8-211">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-211">
        - Settings</span></span><br><span data-ttu-id="26bb8-212">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-212">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-213">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-213">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-214">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-214">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-215">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-215">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-216">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="26bb8-216">Office 2013 on Windows</span></span><br><span data-ttu-id="26bb8-217">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-217">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="26bb8-218">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-218">
        - TaskPane</span></span><br><span data-ttu-id="26bb8-219">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-219">
        - Content</span></span></td>
    <td>  <span data-ttu-id="26bb8-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="26bb8-220">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="26bb8-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-221">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-222">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-222">
        - BindingEvents</span></span><br><span data-ttu-id="26bb8-223">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-223">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-224">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-224">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-225">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-225">
        - File</span></span><br><span data-ttu-id="26bb8-226">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-226">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-227">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-227">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-228">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-228">
        - Selection</span></span><br><span data-ttu-id="26bb8-229">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-229">
        - Settings</span></span><br><span data-ttu-id="26bb8-230">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-230">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-231">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-231">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-232">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-232">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-233">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-233">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-234">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-234">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="26bb8-235">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-235">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="26bb8-236">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-236">- TaskPane</span></span><br><span data-ttu-id="26bb8-237">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-237">
        - Content</span></span></td>
    <td><span data-ttu-id="26bb8-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-238">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-239">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="26bb8-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-240">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="26bb8-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-241">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="26bb8-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-242">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="26bb8-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="26bb8-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="26bb8-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="26bb8-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="26bb8-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-249">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-249">- BindingEvents</span></span><br><span data-ttu-id="26bb8-250">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-250">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-251">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-251">
        - File</span></span><br><span data-ttu-id="26bb8-252">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-252">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-253">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-253">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-254">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-254">
        - Selection</span></span><br><span data-ttu-id="26bb8-255">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-255">
        - Settings</span></span><br><span data-ttu-id="26bb8-256">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-256">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-257">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-257">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-258">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-258">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-259">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-259">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-260">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-260">Office apps on Mac</span></span><br><span data-ttu-id="26bb8-261">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-261">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="26bb8-262">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-262">- TaskPane</span></span><br><span data-ttu-id="26bb8-263">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-263">
        - Content</span></span><br><span data-ttu-id="26bb8-264">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-264">
        - Custom Functions</span></span><br><span data-ttu-id="26bb8-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-265">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="26bb8-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-266">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-267">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="26bb8-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-268">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="26bb8-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-269">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="26bb8-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-270">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="26bb8-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="26bb8-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-272">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="26bb8-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="26bb8-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="26bb8-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="26bb8-278">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-278">- BindingEvents</span></span><br><span data-ttu-id="26bb8-279">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-279">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-280">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-280">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-281">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-281">
        - File</span></span><br><span data-ttu-id="26bb8-282">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-282">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-283">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-283">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-284">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-284">
        - PdfFile</span></span><br><span data-ttu-id="26bb8-285">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-285">
        - Selection</span></span><br><span data-ttu-id="26bb8-286">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-286">
        - Settings</span></span><br><span data-ttu-id="26bb8-287">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-287">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-288">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-288">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-289">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-289">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-290">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-290">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-291">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-291">Office 2019 for Mac</span></span><br><span data-ttu-id="26bb8-292">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-292">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="26bb8-293">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-293">- TaskPane</span></span><br><span data-ttu-id="26bb8-294">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-294">
        - Content</span></span><br><span data-ttu-id="26bb8-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-295">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="26bb8-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-296">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-297">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="26bb8-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-298">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="26bb8-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-299">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="26bb8-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-300">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="26bb8-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-301">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="26bb8-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-302">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="26bb8-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="26bb8-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-304">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-306">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-306">- BindingEvents</span></span><br><span data-ttu-id="26bb8-307">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-307">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-308">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-308">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-309">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-309">
        - File</span></span><br><span data-ttu-id="26bb8-310">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-310">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-311">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-311">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-312">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-312">
        - PdfFile</span></span><br><span data-ttu-id="26bb8-313">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-313">
        - Selection</span></span><br><span data-ttu-id="26bb8-314">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-314">
        - Settings</span></span><br><span data-ttu-id="26bb8-315">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-315">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-316">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-316">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-317">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-317">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-318">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-318">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-319">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-319">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="26bb8-320">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-320">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="26bb8-321">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-321">- TaskPane</span></span><br><span data-ttu-id="26bb8-322">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-322">
        - Content</span></span></td>
    <td><span data-ttu-id="26bb8-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-323">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="26bb8-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-324">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="26bb8-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-325">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="26bb8-326">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-326">- BindingEvents</span></span><br><span data-ttu-id="26bb8-327">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-327">
        - CompressedFile</span></span><br><span data-ttu-id="26bb8-328">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-328">
        - DocumentEvents</span></span><br><span data-ttu-id="26bb8-329">
        - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-329">
        - File</span></span><br><span data-ttu-id="26bb8-330">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-330">
        - MatrixBindings</span></span><br><span data-ttu-id="26bb8-331">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-331">
        - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-332">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-332">
        - PdfFile</span></span><br><span data-ttu-id="26bb8-333">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-333">
        - Selection</span></span><br><span data-ttu-id="26bb8-334">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-334">
        - Settings</span></span><br><span data-ttu-id="26bb8-335">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-335">
        - TableBindings</span></span><br><span data-ttu-id="26bb8-336">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-336">
        - TableCoercion</span></span><br><span data-ttu-id="26bb8-337">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-337">
        - TextBindings</span></span><br><span data-ttu-id="26bb8-338">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-338">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="26bb8-339">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="26bb8-339">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions"></a><span data-ttu-id="26bb8-340">カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-340">Custom Functions</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="26bb8-341">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-341">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="26bb8-342">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-342">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="26bb8-343">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-343">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="26bb8-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-344"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-345">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-345">Office on the web</span></span></td>
    <td><span data-ttu-id="26bb8-346">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-346">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="26bb8-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-347">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-348">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-348">Office on Windows</span></span><br><span data-ttu-id="26bb8-349">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-349">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="26bb8-350">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-350">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="26bb8-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-351">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-352">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="26bb8-352">Office for Mac</span></span><br><span data-ttu-id="26bb8-353">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="26bb8-353">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="26bb8-354">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="26bb8-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="26bb8-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-355">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="26bb8-356">Outlook</span><span class="sxs-lookup"><span data-stu-id="26bb8-356">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="26bb8-357">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-357">Platform</span></span></th>
    <th><span data-ttu-id="26bb8-358">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-358">Extension points</span></span></th>
    <th><span data-ttu-id="26bb8-359">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-359">API requirement sets</span></span></th>
    <th><span data-ttu-id="26bb8-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-360"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-361">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-361">Office on the web</span></span><br><span data-ttu-id="26bb8-362">(モダン)</span><span class="sxs-lookup"><span data-stu-id="26bb8-362">Modern</span></span></td>
    <td> <span data-ttu-id="26bb8-363">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-363">- Mail Read</span></span><br><span data-ttu-id="26bb8-364">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-364">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-365">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-366">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-367">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-368">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-369">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-370">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-371">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="26bb8-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-372">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="26bb8-373">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-373">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-374">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-374">Office on the web</span></span><br><span data-ttu-id="26bb8-375">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="26bb8-375">Classic</span></span></td>
    <td> <span data-ttu-id="26bb8-376">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-376">- Mail Read</span></span><br><span data-ttu-id="26bb8-377">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-377">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-378">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-379">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-384">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="26bb8-385">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-385">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-386">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-386">Office on Windows</span></span><br><span data-ttu-id="26bb8-387">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-387">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-388">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-388">- Mail Read</span></span><br><span data-ttu-id="26bb8-389">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-389">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-390">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="26bb8-391">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="26bb8-391">
      - Modules</span></span></td>
    <td> <span data-ttu-id="26bb8-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="26bb8-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-398">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="26bb8-399">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-399">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-400">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-400">Office 2019 on Windows</span></span><br><span data-ttu-id="26bb8-401">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-401">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-402">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-402">- Mail Read</span></span><br><span data-ttu-id="26bb8-403">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-403">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="26bb8-405">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="26bb8-405">
      - Modules</span></span></td>
    <td> <span data-ttu-id="26bb8-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-406">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="26bb8-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="26bb8-413">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-413">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-414">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-414">Office 2016 on Windows</span></span><br><span data-ttu-id="26bb8-415">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-415">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-416">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-416">- Mail Read</span></span><br><span data-ttu-id="26bb8-417">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-417">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="26bb8-419">
      - モジュール</span><span class="sxs-lookup"><span data-stu-id="26bb8-419">
      - Modules</span></span></td>
    <td> <span data-ttu-id="26bb8-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-420">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="26bb8-424">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-424">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-425">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="26bb8-425">Office 2013 on Windows</span></span><br><span data-ttu-id="26bb8-426">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-426">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-427">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-427">- Mail Read</span></span><br><span data-ttu-id="26bb8-428">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-428">
      - Mail Compose</span></span></td>
    <td> <span data-ttu-id="26bb8-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-429">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-431">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="26bb8-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-432">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="26bb8-433">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-433">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-434">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-434">Office apps on iOS</span></span><br><span data-ttu-id="26bb8-435">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-435">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-436">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-436">- Mail Read</span></span><br><span data-ttu-id="26bb8-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-437">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-438">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="26bb8-443">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-443">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-444">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-444">Office apps on Mac</span></span><br><span data-ttu-id="26bb8-445">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-445">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-446">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-446">- Mail Read</span></span><br><span data-ttu-id="26bb8-447">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-447">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-448">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-449">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-450">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-451">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="26bb8-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="26bb8-456">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-456">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-457">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-457">Office 2019 for Mac</span></span><br><span data-ttu-id="26bb8-458">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-458">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-459">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-459">- Mail Read</span></span><br><span data-ttu-id="26bb8-460">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-460">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-461">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-462">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="26bb8-468">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-468">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-469">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-469">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="26bb8-470">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-470">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-471">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-471">- Mail Read</span></span><br><span data-ttu-id="26bb8-472">
      - メールの作成</span><span class="sxs-lookup"><span data-stu-id="26bb8-472">
      - Mail Compose</span></span><br><span data-ttu-id="26bb8-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-474">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="26bb8-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="26bb8-480">使用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-480">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-481">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-481">Office apps on Android</span></span><br><span data-ttu-id="26bb8-482">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-482">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-483">- メールの読み取り</span><span class="sxs-lookup"><span data-stu-id="26bb8-483">- Mail Read</span></span><br><span data-ttu-id="26bb8-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-484">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-485">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="26bb8-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-486">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="26bb8-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="26bb8-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="26bb8-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="26bb8-490">利用不可</span><span class="sxs-lookup"><span data-stu-id="26bb8-490">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="26bb8-491">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="26bb8-491">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="26bb8-492">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="26bb8-492">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="26bb8-493">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="26bb8-493">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="26bb8-494">Word</span><span class="sxs-lookup"><span data-stu-id="26bb8-494">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="26bb8-495">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-495">Platform</span></span></th>
    <th><span data-ttu-id="26bb8-496">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-496">Extension points</span></span></th>
    <th><span data-ttu-id="26bb8-497">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-497">API requirement sets</span></span></th>
    <th><span data-ttu-id="26bb8-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-498"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-499">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-499">Office on the web</span></span></td>
    <td> <span data-ttu-id="26bb8-500">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-500">- TaskPane</span></span><br><span data-ttu-id="26bb8-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-501">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-502">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-503">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="26bb8-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-504">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="26bb8-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-505">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-506">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-507">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="26bb8-508">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-508">- BindingEvents</span></span><br><span data-ttu-id="26bb8-509">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-509">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-510">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-510">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-511">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-511">
         - File</span></span><br><span data-ttu-id="26bb8-512">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-512">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-513">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-513">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-514">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-514">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-515">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-515">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-516">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-516">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-517">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-517">
         - Selection</span></span><br><span data-ttu-id="26bb8-518">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-518">
         - Settings</span></span><br><span data-ttu-id="26bb8-519">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-519">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-520">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-520">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-521">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-521">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-522">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-522">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-523">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-523">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-524">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-524">Office on Windows</span></span><br><span data-ttu-id="26bb8-525">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-525">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-526">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-526">- TaskPane</span></span><br><span data-ttu-id="26bb8-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-527">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-528">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-529">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="26bb8-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-530">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="26bb8-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-531">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-532">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-533">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="26bb8-534">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-534">- BindingEvents</span></span><br><span data-ttu-id="26bb8-535">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-535">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-536">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-536">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-537">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-537">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-538">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-538">
         - File</span></span><br><span data-ttu-id="26bb8-539">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-539">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-540">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-540">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-541">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-541">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-542">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-542">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-543">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-543">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-544">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-544">
         - Selection</span></span><br><span data-ttu-id="26bb8-545">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-545">
         - Settings</span></span><br><span data-ttu-id="26bb8-546">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-546">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-547">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-547">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-548">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-548">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-549">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-549">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-550">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-550">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-551">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-551">Office 2019 on Windows</span></span><br><span data-ttu-id="26bb8-552">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-552">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-553">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-553">- TaskPane</span></span><br><span data-ttu-id="26bb8-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-554">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-555">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-556">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="26bb8-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-557">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="26bb8-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-558">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-559">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-560">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-560">- BindingEvents</span></span><br><span data-ttu-id="26bb8-561">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-561">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-562">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-562">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-563">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-563">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-564">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-564">
         - File</span></span><br><span data-ttu-id="26bb8-565">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-565">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-566">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-566">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-567">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-567">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-568">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-568">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-569">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-569">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-570">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-570">
         - Selection</span></span><br><span data-ttu-id="26bb8-571">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-571">
         - Settings</span></span><br><span data-ttu-id="26bb8-572">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-572">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-573">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-573">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-574">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-574">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-575">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-575">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-576">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-576">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-577">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-577">Office 2016 on Windows</span></span><br><span data-ttu-id="26bb8-578">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-578">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-579">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-579">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-580">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-581">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="26bb8-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-582">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-583">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-583">- BindingEvents</span></span><br><span data-ttu-id="26bb8-584">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-584">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-585">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-585">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-586">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-586">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-587">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-587">
         - File</span></span><br><span data-ttu-id="26bb8-588">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-588">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-589">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-589">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-590">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-590">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-591">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-591">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-592">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-592">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-593">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-593">
         - Selection</span></span><br><span data-ttu-id="26bb8-594">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-594">
         - Settings</span></span><br><span data-ttu-id="26bb8-595">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-595">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-596">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-596">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-597">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-597">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-598">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-598">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-599">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-599">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-600">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="26bb8-600">Office 2013 on Windows</span></span><br><span data-ttu-id="26bb8-601">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-601">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-602">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-602">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="26bb8-603">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="26bb8-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-604">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-605">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-605">- BindingEvents</span></span><br><span data-ttu-id="26bb8-606">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-606">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-607">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-607">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-608">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-608">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-609">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-609">
         - File</span></span><br><span data-ttu-id="26bb8-610">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-610">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-611">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-611">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-612">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-612">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-613">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-613">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-614">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-614">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-615">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-615">
         - Selection</span></span><br><span data-ttu-id="26bb8-616">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-616">
         - Settings</span></span><br><span data-ttu-id="26bb8-617">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-617">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-618">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-618">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-619">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-619">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-620">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-620">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-621">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-621">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-622">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-622">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="26bb8-623">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-623">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-624">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-624">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-625">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-626">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="26bb8-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-627">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="26bb8-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-628">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-629">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="26bb8-630">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-630">- BindingEvents</span></span><br><span data-ttu-id="26bb8-631">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-631">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-632">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-632">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-633">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-633">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-634">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-634">
         - File</span></span><br><span data-ttu-id="26bb8-635">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-635">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-636">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-636">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-637">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-637">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-638">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-638">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-639">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-639">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-640">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-640">
         - Selection</span></span><br><span data-ttu-id="26bb8-641">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-641">
         - Settings</span></span><br><span data-ttu-id="26bb8-642">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-642">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-643">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-643">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-644">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-644">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-645">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-645">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-646">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-646">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-647">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-647">Office apps on Mac</span></span><br><span data-ttu-id="26bb8-648">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-648">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-649">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-649">- TaskPane</span></span><br><span data-ttu-id="26bb8-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-650">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-651">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-652">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="26bb8-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-653">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="26bb8-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-654">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-655">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-656">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="26bb8-657">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-657">- BindingEvents</span></span><br><span data-ttu-id="26bb8-658">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-658">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-659">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-659">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-660">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-660">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-661">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-661">
         - File</span></span><br><span data-ttu-id="26bb8-662">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-662">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-663">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-663">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-664">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-664">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-665">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-665">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-666">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-666">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-667">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-667">
         - Selection</span></span><br><span data-ttu-id="26bb8-668">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-668">
         - Settings</span></span><br><span data-ttu-id="26bb8-669">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-669">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-670">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-670">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-671">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-671">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-672">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-672">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-673">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-673">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-674">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-674">Office 2019 for Mac</span></span><br><span data-ttu-id="26bb8-675">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-675">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-676">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-676">- TaskPane</span></span><br><span data-ttu-id="26bb8-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-677">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-678">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-679">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="26bb8-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-680">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="26bb8-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-681">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="26bb8-683">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-683">- BindingEvents</span></span><br><span data-ttu-id="26bb8-684">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-684">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-685">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-685">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-686">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-686">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-687">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-687">
         - File</span></span><br><span data-ttu-id="26bb8-688">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-688">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-689">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-689">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-690">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-690">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-691">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-691">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-692">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-692">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-693">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-693">
         - Selection</span></span><br><span data-ttu-id="26bb8-694">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-694">
         - Settings</span></span><br><span data-ttu-id="26bb8-695">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-695">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-696">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-696">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-697">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-697">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-698">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-698">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-699">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-699">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-700">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-700">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="26bb8-701">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-701">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-702">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-702">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-703">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="26bb8-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="26bb8-704">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="26bb8-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-705">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-706">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-706">- BindingEvents</span></span><br><span data-ttu-id="26bb8-707">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-707">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-708">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="26bb8-708">
         - CustomXmlParts</span></span><br><span data-ttu-id="26bb8-709">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-709">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-710">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-710">
         - File</span></span><br><span data-ttu-id="26bb8-711">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-711">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-712">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-712">
         - MatrixBindings</span></span><br><span data-ttu-id="26bb8-713">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-713">
         - MatrixCoercion</span></span><br><span data-ttu-id="26bb8-714">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-714">
         - OoxmlCoercion</span></span><br><span data-ttu-id="26bb8-715">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-715">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-716">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-716">
         - Selection</span></span><br><span data-ttu-id="26bb8-717">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-717">
         - Settings</span></span><br><span data-ttu-id="26bb8-718">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-718">
         - TableBindings</span></span><br><span data-ttu-id="26bb8-719">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-719">
         - TableCoercion</span></span><br><span data-ttu-id="26bb8-720">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="26bb8-720">
         - TextBindings</span></span><br><span data-ttu-id="26bb8-721">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-721">
         - TextCoercion</span></span><br><span data-ttu-id="26bb8-722">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-722">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="26bb8-723">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="26bb8-723">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="26bb8-724">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="26bb8-724">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="26bb8-725">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-725">Platform</span></span></th>
    <th><span data-ttu-id="26bb8-726">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-726">Extension points</span></span></th>
    <th><span data-ttu-id="26bb8-727">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-727">API requirement sets</span></span></th>
    <th><span data-ttu-id="26bb8-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-728"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-729">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-729">Office on the web</span></span></td>
    <td> <span data-ttu-id="26bb8-730">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-730">- Content</span></span><br><span data-ttu-id="26bb8-731">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-731">
         - TaskPane</span></span><br><span data-ttu-id="26bb8-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-732">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-733">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="26bb8-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-734">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-735">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-736">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="26bb8-737">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-737">- ActiveView</span></span><br><span data-ttu-id="26bb8-738">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-738">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-739">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-740">
         - File</span></span><br><span data-ttu-id="26bb8-741">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-741">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-742">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-742">
         - Selection</span></span><br><span data-ttu-id="26bb8-743">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-743">
         - Settings</span></span><br><span data-ttu-id="26bb8-744">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-744">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-745">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-745">Office on Windows</span></span><br><span data-ttu-id="26bb8-746">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-746">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-747">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-747">- Content</span></span><br><span data-ttu-id="26bb8-748">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-748">
         - TaskPane</span></span><br><span data-ttu-id="26bb8-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-749">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-750">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="26bb8-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-751">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-752">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-753">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="26bb8-754">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-754">- ActiveView</span></span><br><span data-ttu-id="26bb8-755">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-755">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-756">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-756">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-757">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-757">
         - File</span></span><br><span data-ttu-id="26bb8-758">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-758">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-759">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-759">
         - Selection</span></span><br><span data-ttu-id="26bb8-760">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-760">
         - Settings</span></span><br><span data-ttu-id="26bb8-761">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-761">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-762">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-762">Office 2019 on Windows</span></span><br><span data-ttu-id="26bb8-763">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-763">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-764">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-764">- Content</span></span><br><span data-ttu-id="26bb8-765">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-765">
         - TaskPane</span></span><br><span data-ttu-id="26bb8-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-767">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-768">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-769">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-769">- ActiveView</span></span><br><span data-ttu-id="26bb8-770">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-770">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-771">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-771">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-772">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-772">
         - File</span></span><br><span data-ttu-id="26bb8-773">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-773">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-774">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-774">
         - Selection</span></span><br><span data-ttu-id="26bb8-775">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-775">
         - Settings</span></span><br><span data-ttu-id="26bb8-776">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-776">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-777">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-777">Office 2016 on Windows</span></span><br><span data-ttu-id="26bb8-778">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-778">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-779">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-779">- Content</span></span><br><span data-ttu-id="26bb8-780">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-780">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="26bb8-781">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="26bb8-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-783">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-783">- ActiveView</span></span><br><span data-ttu-id="26bb8-784">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-784">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-785">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-785">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-786">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-786">
         - File</span></span><br><span data-ttu-id="26bb8-787">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-787">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-788">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-788">
         - Selection</span></span><br><span data-ttu-id="26bb8-789">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-789">
         - Settings</span></span><br><span data-ttu-id="26bb8-790">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-790">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-791">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="26bb8-791">Office 2013 on Windows</span></span><br><span data-ttu-id="26bb8-792">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-792">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-793">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-793">- Content</span></span><br><span data-ttu-id="26bb8-794">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-794">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="26bb8-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="26bb8-795">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="26bb8-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-797">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-797">- ActiveView</span></span><br><span data-ttu-id="26bb8-798">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-798">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-799">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-799">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-800">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-800">
         - File</span></span><br><span data-ttu-id="26bb8-801">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-801">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-802">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-802">
         - Selection</span></span><br><span data-ttu-id="26bb8-803">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-803">
         - Settings</span></span><br><span data-ttu-id="26bb8-804">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-804">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-805">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-805">Sideload Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="26bb8-806">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-806">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-807">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-807">- Content</span></span><br><span data-ttu-id="26bb8-808">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-808">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-809">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="26bb8-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-810">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-811">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-812">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-812">- ActiveView</span></span><br><span data-ttu-id="26bb8-813">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-813">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-814">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-814">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-815">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-815">
         - File</span></span><br><span data-ttu-id="26bb8-816">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-816">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-817">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-817">
         - Selection</span></span><br><span data-ttu-id="26bb8-818">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-818">
         - Settings</span></span><br><span data-ttu-id="26bb8-819">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-819">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-820">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="26bb8-820">Office apps on Mac</span></span><br><span data-ttu-id="26bb8-821">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="26bb8-821">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="26bb8-822">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-822">- Content</span></span><br><span data-ttu-id="26bb8-823">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-823">
         - TaskPane</span></span><br><span data-ttu-id="26bb8-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-824">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-825">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="26bb8-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-827">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="26bb8-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-828">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="26bb8-829">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-829">- ActiveView</span></span><br><span data-ttu-id="26bb8-830">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-830">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-831">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-831">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-832">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-832">
         - File</span></span><br><span data-ttu-id="26bb8-833">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-833">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-834">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-834">
         - Selection</span></span><br><span data-ttu-id="26bb8-835">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-835">
         - Settings</span></span><br><span data-ttu-id="26bb8-836">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-836">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-837">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-837">Office 2019 for Mac</span></span><br><span data-ttu-id="26bb8-838">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-838">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-839">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-839">- Content</span></span><br><span data-ttu-id="26bb8-840">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-840">
         - TaskPane</span></span><br><span data-ttu-id="26bb8-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-842">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-843">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-844">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-844">- ActiveView</span></span><br><span data-ttu-id="26bb8-845">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-845">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-846">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-846">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-847">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-847">
         - File</span></span><br><span data-ttu-id="26bb8-848">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-848">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-849">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-849">
         - Selection</span></span><br><span data-ttu-id="26bb8-850">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-850">
         - Settings</span></span><br><span data-ttu-id="26bb8-851">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-851">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-852">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-852">Activate Office 2016 on Mac</span></span><br><span data-ttu-id="26bb8-853">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-853">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-854">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-854">- Content</span></span><br><span data-ttu-id="26bb8-855">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-855">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="26bb8-856">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="26bb8-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-858">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="26bb8-858">- ActiveView</span></span><br><span data-ttu-id="26bb8-859">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-859">
         - CompressedFile</span></span><br><span data-ttu-id="26bb8-860">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-860">
         - DocumentEvents</span></span><br><span data-ttu-id="26bb8-861">
         - File</span><span class="sxs-lookup"><span data-stu-id="26bb8-861">
         - File</span></span><br><span data-ttu-id="26bb8-862">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="26bb8-862">
         - PdfFile</span></span><br><span data-ttu-id="26bb8-863">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-863">
         - Selection</span></span><br><span data-ttu-id="26bb8-864">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-864">
         - Settings</span></span><br><span data-ttu-id="26bb8-865">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-865">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="26bb8-866">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="26bb8-866">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="26bb8-867">OneNote</span><span class="sxs-lookup"><span data-stu-id="26bb8-867">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="26bb8-868">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-868">Platform</span></span></th>
    <th><span data-ttu-id="26bb8-869">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-869">Extension points</span></span></th>
    <th><span data-ttu-id="26bb8-870">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-870">API requirement sets</span></span></th>
    <th><span data-ttu-id="26bb8-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-871"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-872">Office on the web</span><span class="sxs-lookup"><span data-stu-id="26bb8-872">Office on the web</span></span></td>
    <td> <span data-ttu-id="26bb8-873">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="26bb8-873">- Content</span></span><br><span data-ttu-id="26bb8-874">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-874">
         - TaskPane</span></span><br><span data-ttu-id="26bb8-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-875">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="26bb8-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-876">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="26bb8-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-877">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="26bb8-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-878">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-879">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="26bb8-879">- DocumentEvents</span></span><br><span data-ttu-id="26bb8-880">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-880">
         - HtmlCoercion</span></span><br><span data-ttu-id="26bb8-881">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="26bb8-881">
         - Settings</span></span><br><span data-ttu-id="26bb8-882">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-882">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="26bb8-883">Project</span><span class="sxs-lookup"><span data-stu-id="26bb8-883">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="26bb8-884">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="26bb8-884">Platform</span></span></th>
    <th><span data-ttu-id="26bb8-885">拡張点</span><span class="sxs-lookup"><span data-stu-id="26bb8-885">Extension points</span></span></th>
    <th><span data-ttu-id="26bb8-886">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-886">API requirement sets</span></span></th>
    <th><span data-ttu-id="26bb8-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="26bb8-887"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-888">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="26bb8-888">Office 2019 on Windows</span></span><br><span data-ttu-id="26bb8-889">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-889">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-890">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-890">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-891">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-892">- Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-892">- Selection</span></span><br><span data-ttu-id="26bb8-893">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-893">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-894">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="26bb8-894">Office 2016 on Windows</span></span><br><span data-ttu-id="26bb8-895">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-895">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-896">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-896">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-897">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-898">- Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-898">- Selection</span></span><br><span data-ttu-id="26bb8-899">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-899">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="26bb8-900">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="26bb8-900">Office 2013 on Windows</span></span><br><span data-ttu-id="26bb8-901">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="26bb8-901">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="26bb8-902">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="26bb8-902">- TaskPane</span></span></td>
    <td> <span data-ttu-id="26bb8-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="26bb8-903">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="26bb8-904">- Selection</span><span class="sxs-lookup"><span data-stu-id="26bb8-904">- Selection</span></span><br><span data-ttu-id="26bb8-905">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="26bb8-905">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="26bb8-906">関連項目</span><span class="sxs-lookup"><span data-stu-id="26bb8-906">See also</span></span>

- [<span data-ttu-id="26bb8-907">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="26bb8-907">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="26bb8-908">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-908">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="26bb8-909">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-909">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="26bb8-910">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="26bb8-910">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="26bb8-911">JavaScript API for Office リファレンス</span><span class="sxs-lookup"><span data-stu-id="26bb8-911">JavaScript API for Office reference</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="26bb8-912">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="26bb8-912">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="26bb8-913">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="26bb8-913">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="26bb8-914">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="26bb8-914">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="26bb8-915">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="26bb8-915">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="26bb8-916">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="26bb8-916">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="26bb8-917">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="26bb8-917">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
